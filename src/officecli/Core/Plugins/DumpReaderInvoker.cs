// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using OfficeCli.Handlers;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// Runs a dump-reader plugin per docs/plugin-protocol.md §5.1. The plugin
/// reads a foreign source file (e.g. .doc) and **streams** BatchItem objects
/// as JSONL (one JSON object per line) on stdout. Main opens a fresh native
/// scratch file (extension from the plugin's manifest <c>target</c> field —
/// docx/xlsx/pptx), replays each item as it arrives, and returns the
/// populated path.
///
/// Streaming has two benefits over a buffered JSON-array transport: per-line
/// activity feeds the idle watchdog (§5.6), and main's memory does not scale
/// with batch size on multi-million-paragraph .doc files.
///
/// The conversion is one-shot: edits to the returned file are not propagated
/// back to the source file.
/// </summary>
public static class DumpReaderInvoker
{
    public sealed record DumpResult(string ConvertedPath, ResolvedPlugin Plugin);

    /// <summary>
    /// Resolve a dump-reader plugin for <paramref name="sourceExt"/>, invoke it
    /// against <paramref name="sourceFullPath"/>, and replay the resulting
    /// JSONL stream into a fresh native file. Throws CliException on
    /// resolution or invocation failure; otherwise the result references a
    /// temp file the caller must dispose (or leave for OS tmp cleanup).
    /// </summary>
    /// <summary>
    /// Warnings produced during sibling-cache generation that the caller (e.g.
    /// ResidentServer) should surface to the user on the next command's stderr.
    /// The handler-open path runs at resident-startup time, before any per
    /// command Console.SetError scope is in place — Console.Error.WriteLine
    /// there is captured by an unread pipe and lost. Drain this list inside a
    /// captured stderr scope to deliver the message.
    /// </summary>
    public static readonly Queue<string> PendingWarnings = new();
    private static readonly object _pendingLock = new();

    /// <summary>
    /// Emit any warnings queued since the last drain to <see cref="Console.Error"/>.
    /// Called by ResidentServer on each command boundary (where Console.Error is
    /// captured into the response envelope), and by the direct path before
    /// returning to the user.
    /// </summary>
    public static void DrainPendingWarnings()
    {
        lock (_pendingLock)
        {
            while (PendingWarnings.Count > 0)
                Console.Error.WriteLine(PendingWarnings.Dequeue());
        }
    }

    private static void QueueWarning(string text)
    {
        lock (_pendingLock) PendingWarnings.Enqueue(text);
    }

    public static DumpResult Run(string sourceFullPath, string sourceExt)
    {
        var plugin = PluginRegistry.FindFor(PluginKind.DumpReader, sourceExt)
            ?? throw new CliException($"No dump-reader plugin found for {sourceExt}.")
            {
                Code = "dump_reader_not_found",
                Suggestion = "Install a dump-reader plugin (`officecli plugins list` to see installed; docs/plugin-protocol.md for paths).",
            };

        var targetExt = plugin.Manifest.ResolveTargetExtension();
        var tmpOut = Path.Combine(Path.GetTempPath(),
            $"officecli-dumpread-{Guid.NewGuid():N}{targetExt}");
        // minimal: true gives a bare-skeleton native file (no default styles,
        // theme, or docDefaults for docx; equivalent skeleton for xlsx/pptx).
        // The plugin's batch is expected to define everything it references —
        // round-trip dumps from `officecli dump` do exactly that.
        BlankDocCreator.Create(tmpOut, locale: null, minimal: true);

        int itemIndex = 0;
        Exception? replayError = null;

        try
        {
            using var handler = DocumentHandlerFactory.Open(tmpOut, editable: true);

            void OnLine(string raw)
            {
                // Strip a per-line UTF-8 BOM (U+FEFF). Some Windows JSON
                // serializers emit BOM on every line of JSONL output, which
                // is technically RFC 8259 noncompliant but easy to absorb at
                // the host. Trim handles trailing whitespace and CR from a
                // CRLF-on-Windows plugin.
                var line = raw.TrimStart('﻿').Trim();
                if (line.Length == 0) return;

                // Reject legacy top-level JSON arrays explicitly. Plugins that
                // emitted `[...]` under the old protocol now fail with a clear
                // error instead of being parsed as a malformed BatchItem.
                if (line[0] == '[')
                    throw new CliException(
                        $"Dump-reader plugin '{plugin.Manifest.Name}' emitted a JSON array; protocol v1 requires JSONL (one BatchItem per line).")
                    { Code = "corrupt_batch" };

                BatchItem? item;
                try
                {
                    item = JsonSerializer.Deserialize(line, BatchJsonContext.Default.BatchItem);
                }
                catch (JsonException ex)
                {
                    throw new CliException(
                        $"Dump-reader plugin '{plugin.Manifest.Name}' emitted invalid JSON at item #{itemIndex}: {ex.Message}")
                    { Code = "plugin_contract_violation" };
                }
                if (item is null)
                    throw new CliException(
                        $"Dump-reader plugin '{plugin.Manifest.Name}' emitted null at item #{itemIndex}.")
                    { Code = "plugin_contract_violation" };

                try
                {
                    CommandBuilder.ExecuteBatchItem(handler, item, json: false);
                }
                catch (Exception ex)
                {
                    throw new CliException(
                        $"Dump-reader plugin '{plugin.Manifest.Name}' command #{itemIndex} ({item.Command}) failed while replaying: {ex.Message}", ex)
                    { Code = "plugin_command_failed" };
                }

                itemIndex++;
            }

            var idle = plugin.Manifest.ResolveIdleTimeout("dump");
            var result = PluginProcess.Run(new PluginProcess.RunOptions
            {
                ExecutablePath = plugin.ExecutablePath,
                Arguments = new[] { "dump", sourceFullPath },
                IdleTimeoutSeconds = idle,
                OnStdoutLine = OnLine,
            });

            // Bubble up the per-line callback error first — its message is more
            // actionable than the generic non-zero exit that follows.
            if (PluginProcess.LineCallbackError is CliException ce)
                throw ce;
            if (PluginProcess.LineCallbackError is not null)
                throw new CliException(
                    $"Dump-reader plugin '{plugin.Manifest.Name}' replay aborted: {PluginProcess.LineCallbackError.Message}",
                    PluginProcess.LineCallbackError)
                { Code = "plugin_command_failed" };

            if (result.IdleTimedOut)
                throw new CliException(
                    $"Dump-reader plugin '{plugin.Manifest.Name}' produced no output for {idle}s — likely hung.")
                {
                    Code = "plugin_idle_timeout",
                    Suggestion = $"Override with --timeout 0 or set a longer `idle_timeout_seconds.verbs.dump` in the plugin's manifest.",
                };

            if (result.ExitCode != 0)
                throw new CliException(
                    $"Dump-reader plugin '{plugin.Manifest.Name}' failed (exit {result.ExitCode}): {Truncate(result.Stderr, 500)}")
                {
                    Code = result.ExitCode switch
                    {
                        2 => "corrupt_input",
                        3 => "unsupported_feature",
                        4 => "license_expired",
                        5 => "protocol_mismatch",
                        6 => "plugin_idle_timeout",
                        _ => "plugin_failed",
                    },
                };

            // Empty output + exit 0 is ambiguous: the .doc might genuinely be
            // blank, or the plugin might have silently skipped content it does
            // not yet know how to translate. Queue a warning that the host
            // will surface on the next command boundary (ResidentServer drains
            // PendingWarnings inside its per-command stderr scope, so the
            // message reaches the user even though the dump runs at handler
            // open time — outside any captured-stderr scope).
            if (itemIndex == 0)
                QueueWarning(
                    $"[warning] dump-reader plugin '{plugin.Manifest.Name}' produced no commands for {Path.GetFileName(sourceFullPath)}. " +
                    $"The generated {targetExt} will be blank — this is usually a plugin gap, not a source-file property.");
        }
        catch (Exception ex)
        {
            replayError = ex;
            try { File.Delete(tmpOut); } catch { }
            throw;
        }
        finally
        {
            // If we threw, the tmp file is already cleaned up above. If we
            // succeeded, the caller takes ownership. Either way, leave nothing
            // dangling on disk.
            if (replayError is null && !File.Exists(tmpOut))
                throw new CliException(
                    $"Dump-reader plugin '{plugin.Manifest.Name}' replay produced no output file.")
                { Code = "plugin_contract_violation" };
        }

        return new DumpResult(tmpOut, plugin);
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";
}
