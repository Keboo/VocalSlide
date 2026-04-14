using Keboo.VocalSlide.Services;

namespace Keboo.VocalSlide.Tests;

public class PresenterNotesParserTests
{
    [Test]
    public async Task Parse_WhenDelimiterExists_SplitsPresenterNotesFromAutomationPrompt()
    {
        PresenterNotesParser parser = new();

        var result = parser.Parse(
            """
            Human note line one
            Human note line two
            ---
            Switch when I begin talking about quarterly revenue and the updated forecast.
            """);

        await Assert.That(result.PresenterNotes).IsEqualTo("Human note line one\nHuman note line two");
        await Assert.That(result.AutomationPrompt).IsEqualTo("Switch when I begin talking about quarterly revenue and the updated forecast.");
        await Assert.That(result.HasDelimiter).IsTrue();
    }

    [Test]
    public async Task Parse_WhenDelimiterIsMissing_KeepsOnlyPresenterNotes()
    {
        PresenterNotesParser parser = new();

        var result = parser.Parse("These notes never define an automation prompt.");

        await Assert.That(result.PresenterNotes).IsEqualTo("These notes never define an automation prompt.");
        await Assert.That(result.AutomationPrompt).IsEqualTo(string.Empty);
        await Assert.That(result.HasDelimiter).IsFalse();
    }

    [Test]
    public async Task Parse_WhenDelimiterUsesMoreThanThreeHyphens_StillSplitsTheSections()
    {
        PresenterNotesParser parser = new();

        var result = parser.Parse(
            """
            Prompt the audience with a question first.
            ------
            Switch when I start outlining the migration timeline.
            """);

        await Assert.That(result.PresenterNotes).IsEqualTo("Prompt the audience with a question first.");
        await Assert.That(result.AutomationPrompt).IsEqualTo("Switch when I start outlining the migration timeline.");
        await Assert.That(result.HasDelimiter).IsTrue();
    }

    [Test]
    public async Task Parse_WhenPromptStartsImmediatelyAfterDelimiter_ReturnsEmptyPresenterNotes()
    {
        PresenterNotesParser parser = new();

        var result = parser.Parse(
            """
            ---
            This is about Maddy
            """);

        await Assert.That(result.PresenterNotes).IsEqualTo(string.Empty);
        await Assert.That(result.AutomationPrompt).IsEqualTo("This is about Maddy");
        await Assert.That(result.HasDelimiter).IsTrue();
    }

    [Test]
    public async Task Parse_WhenPowerPointReturnsCarriageReturns_StillSplitsSections()
    {
        PresenterNotesParser parser = new();

        var result = parser.Parse("---\rThis is about Maddy");

        await Assert.That(result.PresenterNotes).IsEqualTo(string.Empty);
        await Assert.That(result.AutomationPrompt).IsEqualTo("This is about Maddy");
        await Assert.That(result.HasDelimiter).IsTrue();
    }
}
