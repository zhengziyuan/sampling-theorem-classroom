param([string]$SourceWav = "", [string]$OutputPptx = "", [string]$PythonExe = "")
# Compatibility entry point for the command printed in the original audio deck.
$taskAudioBuilder = Join-Path $PSScriptRoot "..\..\tools\audio\build_audio_sampling_demo.ps1"
& $taskAudioBuilder @PSBoundParameters
