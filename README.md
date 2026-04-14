# VocalSlide

VocalSlide is a Windows WPF companion app that listens to your microphone, transcribes speech locally, evaluates indexed PowerPoint slide prompts across the whole deck with a local GGUF LLM, and moves the deck for you during a live presentation.

## Presenter notes format

Each slide should keep the real speaker notes above a delimiter of at least three hyphens, and the machine prompt below it.

```text
Remind the room why the old plan missed the mark.
Call out the revised timeline before showing the chart.
---
Switch when I start talking about the failed rollout, the updated timeline, and the chart that explains the recovery plan.
```

## Runtime requirements

- Windows desktop PowerPoint installed locally
- An open PowerPoint presentation, with the slide show running
- Either local model files already on disk, or internet access so the app can download them into `%LOCALAPPDATA%\Keboo.VocalSlide`

## Using the app

1. Open the PowerPoint deck and start the slide show.
2. Use **Download Whisper** to fetch the recommended `ggml-tiny.en.bin` transcription model, or select the larger base model if you want a little more accuracy.
3. Use **Download GGUF** to fetch the default `Qwen2.5-0.5B-Instruct` GGUF file, or replace the URL with any other direct GGUF download URL you want the app to store locally.
4. Click **Refresh PowerPoint** to load slides and parsed prompts.
5. Click **Start Listening** to begin microphone capture and local evaluation.
6. Use **Previous Slide**, **Next Slide**, or select a row in the slide grid and click **Go To Selected Slide** whenever you want to reposition manually.
7. Leave **Auto-advance enabled** checked if you want the app to move slides automatically, or uncheck it to keep transcription running without changing slides.

The transcript window is now **slide-scoped**: whenever the active slide changes, earlier transcript is cleared so the next evaluation only considers speech captured after that slide change.

Automation keeps a running index of every slide that has an automation prompt and can jump **backward or forward** to whichever indexed slide best matches the current spoken content.

Downloaded models are stored in `%LOCALAPPDATA%\Keboo.VocalSlide`, and the app keeps the active Whisper/GGUF paths pointed at that folder by default.

## Development

Run the solution commands from the `Keboo.VocalSlide` directory so `dotnet` picks up the test-runner settings in `global.json`.

```powershell
Set-Location .\Keboo.VocalSlide
dotnet build .\Keboo.VocalSlide.slnx
dotnet test --solution .\Keboo.VocalSlide.slnx
```
