// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Help;

namespace OfficeCli;

static partial class CommandBuilder
{
    // Recognized verbs that route help through the operation-scoped filter.
    // Matches IDocumentHandler's public surface — keep in sync if new verbs
    // are added to the handler API.
    private static readonly string[] HelpVerbs =
        { "add", "set", "get", "query", "remove" };

    /// <summary>
    /// `officecli help [format] [verb] [element] [--json]` — schema-driven help.
    ///
    /// Argument forms accepted:
    ///   help                         → list formats
    ///   help &lt;format&gt;                → list all elements
    ///   help &lt;format&gt; &lt;verb&gt;         → list elements supporting that verb
    ///   help &lt;format&gt; &lt;element&gt;      → full element detail
    ///   help &lt;format&gt; &lt;verb&gt; &lt;element&gt; → verb-filtered element detail
    ///
    /// The middle arg is interpreted as verb iff it matches HelpVerbs.
    /// Mirrors the actual CLI structure: `officecli &lt;verb&gt; &lt;file&gt; ...`, so
    /// `officecli help docx add chart` reads exactly like the command you
    /// are about to run.
    /// </summary>
    public static Command BuildHelpCommand(Option<bool> jsonOption)
    {
        var formatArg = new Argument<string?>("format")
        {
            Description = "Document format: docx/xlsx/pptx (aliases: word, excel, ppt, powerpoint). Omit to list formats.",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var secondArg = new Argument<string?>("verb-or-element")
        {
            Description = "Verb (add/set/get/query/remove) or element name. Omit to list all elements.",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var thirdArg = new Argument<string?>("element")
        {
            Description = "Element name when a verb was given (e.g. 'help docx add chart').",
            Arity = ArgumentArity.ZeroOrOne,
        };

        var command = new Command("help", "Show schema-driven capability reference for officecli.");
        command.Add(formatArg);
        command.Add(secondArg);
        command.Add(thirdArg);
        command.Add(jsonOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            var format = result.GetValue(formatArg);
            var second = result.GetValue(secondArg);
            var third = result.GetValue(thirdArg);

            // Disambiguate middle arg: is it a verb or an element?
            string? verb = null;
            string? element = null;
            if (second != null)
            {
                if (third != null)
                {
                    // 3 args: format, verb, element — second MUST be a verb.
                    verb = second;
                    element = third;
                }
                else if (HelpVerbs.Contains(second, StringComparer.OrdinalIgnoreCase))
                {
                    // 2 args where second is a verb: filter listing by verb.
                    verb = second;
                }
                else
                {
                    // 2 args where second is NOT a verb: treat as element.
                    element = second;
                }
            }

            return SafeRun(() => RunHelp(format, verb, element, json), json);
        });

        return command;
    }

    private static int RunHelp(string? format, string? verb, string? element, bool json)
    {
        // Case 1: no args — list formats and usage banner.
        if (string.IsNullOrEmpty(format))
        {
            Console.WriteLine("officecli help — capability reference (schema-driven)");
            Console.WriteLine();
            Console.WriteLine("Formats:");
            foreach (var f in SchemaHelpLoader.ListFormats())
                Console.WriteLine($"  {f}");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  officecli help <format>                         List all elements");
            Console.WriteLine("  officecli help <format> <verb>                  Elements supporting the verb");
            Console.WriteLine("  officecli help <format> <element>               Full element detail");
            Console.WriteLine("  officecli help <format> <verb> <element>        Verb-filtered element detail");
            Console.WriteLine("  officecli help <format> <element> --json        Raw schema JSON");
            Console.WriteLine();
            Console.WriteLine("Verbs: add, set, get, query, remove");
            Console.WriteLine("Aliases: word→docx, excel→xlsx, ppt/powerpoint→pptx");
            return 0;
        }

        // Validate verb if supplied.
        if (verb != null && !HelpVerbs.Contains(verb, StringComparer.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine($"error: unknown verb '{verb}'. Valid: {string.Join(", ", HelpVerbs)}.");
            return 1;
        }

        var canonicalFormat = SchemaHelpLoader.NormalizeFormat(format);

        // Case 2: format (+ optional verb) only — list elements.
        if (string.IsNullOrEmpty(element))
        {
            var all = SchemaHelpLoader.ListElements(canonicalFormat);
            var filtered = verb == null
                ? all
                : all.Where(el => SchemaHelpLoader.ElementSupportsVerb(canonicalFormat, el, verb!)).ToList();

            if (filtered.Count == 0 && verb != null)
            {
                Console.WriteLine($"No elements in {canonicalFormat} support '{verb}'.");
                return 0;
            }

            var header = verb == null
                ? $"Elements for {canonicalFormat}:"
                : $"Elements for {canonicalFormat} supporting '{verb}':";
            Console.WriteLine(header);
            foreach (var el in filtered)
                Console.WriteLine($"  {el}");
            Console.WriteLine();

            var detailHint = verb == null
                ? $"Run 'officecli help {canonicalFormat} <element>' for detail."
                : $"Run 'officecli help {canonicalFormat} {verb} <element>' for verb-filtered detail.";
            Console.WriteLine(detailHint);
            return 0;
        }

        // Case 3: format + (optional verb) + element — render schema.
        using var doc = SchemaHelpLoader.LoadSchema(format, element);
        Console.WriteLine(json
            ? SchemaHelpRenderer.RenderJson(doc)
            : SchemaHelpRenderer.RenderHuman(doc, verb));
        return 0;
    }
}
