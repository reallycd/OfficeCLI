// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// Phase 6 of the revision (trackChange) redesign — selector + per-target
// accept/reject. Replaces the legacy `set / --prop acceptAllChanges=all`
// magic-property entry with a uniform path-style addressing:
//
//   set <doc> /revision --prop revision=accept                  # all
//   set <doc> '/revision[@author=Alice]' --prop revision=accept # by author
//   set <doc> '/revision[@type=ins]' --prop revision=reject     # by type
//   set <doc> '/revision[@author=Bob][@type=del]' --prop revision=accept
//   set <doc> /revision[3] --prop revision=accept               # by index
//
// All forms start with `/` to satisfy the CLI no-slash-reject guard
// (CONSISTENCY(no-slash-reject) in CommandBuilder.Set.cs) — selector-mode
// inside `set` is otherwise disabled on the user-facing CLI to prevent
// typo-induced doc corruption. The leading slash makes the revision
// dispatch explicit and consistent with `/body/p[N]` etc.
//
// Action values accepted: accept / reject (case-insensitive).
//
// Path-shift safety: matching elements are processed in reverse document
// order. Some accept/reject actions remove sibling content (delete-runs,
// paragraph-merge on ¶ del); walking forward would shift later /revision[N]
// indexes mid-iteration. Reverse traversal keeps each remaining index
// valid until its turn.

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>Single revision marker discovered in document order.</summary>
    private sealed record RevisionRef(
        OpenXmlElement Element,
        string Kind,
        string? Author,
        DateTime? Date,
        string? Id);

    /// <summary>True if <paramref name="path"/> is a revision selector / index
    /// path that the revision dispatcher should claim. Matches bare `revision`,
    /// `revision[...attr-filter...]`, and `/revision[N]` (the indexed form
    /// surfaced by `query revision`).</summary>
    private static bool IsRevisionSelectorPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return false;
        if (path == "/revision") return true;
        if (Regex.IsMatch(path, @"^/revision\[(?:\d+|@\w+=[^\]]+)\](?:\[@\w+=[^\]]+\])*$"))
            return true;
        return false;
    }

    /// <summary>Set entry for revision selectors. Returns the unsupported-key
    /// list (always empty in the happy path — selector dispatch only consumes
    /// `revision`).</summary>
    internal List<string> SetRevisionsBySelector(string path, Dictionary<string, string> properties)
    {
        Modified = true;
        var unsupported = new List<string>();

        if (!properties.TryGetValue("revision", out var action) || string.IsNullOrEmpty(action))
        {
            throw new ArgumentException(
                "revision selector requires --prop revision=accept|reject");
        }
        var act = action.Trim().ToLowerInvariant();
        if (act is not ("accept" or "reject"))
        {
            throw new ArgumentException(
                $"revision must be `accept` or `reject` (got `{action}`)");
        }

        var indexMatch = Regex.Match(path, @"^/revision\[(\d+)\]$");
        Predicate<RevisionRef> filter;
        if (indexMatch.Success)
        {
            var idx = int.Parse(indexMatch.Groups[1].Value);
            var all = EnumerateRevisions();
            if (idx < 1 || idx > all.Count)
                throw new ArgumentException(
                    $"/revision[{idx}] out of range (document has {all.Count} revisions)");
            var picked = all[idx - 1];
            filter = r => ReferenceEquals(r.Element, picked.Element);
        }
        else
        {
            // Parse [@attr=value] segments. Multiple are AND-joined.
            var filterDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match m in Regex.Matches(path, @"\[@(\w+)=([^\]]+)\]"))
            {
                filterDict[m.Groups[1].Value] = m.Groups[2].Value.Trim('"', '\'');
            }
            filter = r => MatchesFilter(r, filterDict);
        }

        var targets = EnumerateRevisions().Where(r => filter(r)).ToList();
        // Reverse so removing earlier siblings doesn't invalidate later refs.
        targets.Reverse();
        foreach (var rev in targets)
        {
            if (act == "accept") AcceptRevision(rev);
            else RejectRevision(rev);
        }
        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    /// <summary>Walk the document body and emit a RevisionRef for every
    /// tracked-change marker. Order is document order (Descendants is
    /// pre-order).</summary>
    private List<RevisionRef> EnumerateRevisions()
    {
        var list = new List<RevisionRef>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return list;

        foreach (var ins in body.Descendants<InsertedRun>())
            list.Add(new RevisionRef(ins, "insertion", ins.Author?.Value, ins.Date?.Value, ins.Id?.Value?.ToString()));
        foreach (var del in body.Descendants<DeletedRun>())
            list.Add(new RevisionRef(del, "deletion", del.Author?.Value, del.Date?.Value, del.Id?.Value?.ToString()));
        foreach (var rpc in body.Descendants<RunPropertiesChange>())
            list.Add(new RevisionRef(rpc, "formatChange", rpc.Author?.Value, rpc.Date?.Value, rpc.Id?.Value?.ToString()));
        foreach (var ppc in body.Descendants<ParagraphPropertiesChange>())
            list.Add(new RevisionRef(ppc, "paragraphChange", ppc.Author?.Value, ppc.Date?.Value, ppc.Id?.Value?.ToString()));
        foreach (var spc in body.Descendants<SectionPropertiesChange>())
            list.Add(new RevisionRef(spc, "sectionChange", spc.Author?.Value, spc.Date?.Value, spc.Id?.Value?.ToString()));
        foreach (var tpc in body.Descendants<TablePropertiesChange>())
            list.Add(new RevisionRef(tpc, "tableChange", tpc.Author?.Value, tpc.Date?.Value, tpc.Id?.Value?.ToString()));
        foreach (var trpc in body.Descendants<TableRowPropertiesChange>())
            list.Add(new RevisionRef(trpc, "rowChange", trpc.Author?.Value, trpc.Date?.Value, trpc.Id?.Value?.ToString()));
        foreach (var tcpc in body.Descendants<TableCellPropertiesChange>())
            list.Add(new RevisionRef(tcpc, "cellChange", tcpc.Author?.Value, tcpc.Date?.Value, tcpc.Id?.Value?.ToString()));
        // trPr/ins, trPr/del — row-level insertion/deletion markers
        foreach (var trPr in body.Descendants<TableRowProperties>())
        {
            var trIns = trPr.GetFirstChild<Inserted>();
            if (trIns != null)
                list.Add(new RevisionRef(trIns, "rowInsertion", trIns.Author?.Value, trIns.Date?.Value, trIns.Id?.Value?.ToString()));
            var trDel = trPr.GetFirstChild<Deleted>();
            if (trDel != null)
                list.Add(new RevisionRef(trDel, "rowDeletion", trDel.Author?.Value, trDel.Date?.Value, trDel.Id?.Value?.ToString()));
        }
        // cellIns, cellDel — cell-level insertion/deletion
        foreach (var tcPr in body.Descendants<TableCellProperties>())
        {
            var ci = tcPr.GetFirstChild<CellInsertion>();
            if (ci != null)
                list.Add(new RevisionRef(ci, "cellInsertion", ci.Author?.Value, ci.Date?.Value, ci.Id?.Value?.ToString()));
            var cd = tcPr.GetFirstChild<CellDeletion>();
            if (cd != null)
                list.Add(new RevisionRef(cd, "cellDeletion", cd.Author?.Value, cd.Date?.Value, cd.Id?.Value?.ToString()));
        }
        // paraMarkIns / paraMarkDel — paragraph mark insertion/deletion
        foreach (var pMark in body.Descendants<ParagraphMarkRunProperties>())
        {
            var pIns = pMark.GetFirstChild<Inserted>();
            if (pIns != null)
                list.Add(new RevisionRef(pIns, "paragraphMarkInsertion", pIns.Author?.Value, pIns.Date?.Value, pIns.Id?.Value?.ToString()));
            var pDel = pMark.GetFirstChild<Deleted>();
            if (pDel != null)
                list.Add(new RevisionRef(pDel, "paragraphMarkDeletion", pDel.Author?.Value, pDel.Date?.Value, pDel.Id?.Value?.ToString()));
        }
        // moveFrom / moveTo
        foreach (var mf in body.Descendants<MoveFromRun>())
            list.Add(new RevisionRef(mf, "moveFrom", mf.Author?.Value, mf.Date?.Value, mf.Id?.Value?.ToString()));
        foreach (var mt in body.Descendants<MoveToRun>())
            list.Add(new RevisionRef(mt, "moveTo", mt.Author?.Value, mt.Date?.Value, mt.Id?.Value?.ToString()));
        return list;
    }

    /// <summary>Match a RevisionRef against the parsed [@attr=value] filter.
    /// Recognised attrs: author, date, id, type. Unknown attrs are rejected
    /// to avoid silent passes (typo in --filter loses safety).</summary>
    private static bool MatchesFilter(RevisionRef rev, Dictionary<string, string> filter)
    {
        foreach (var (key, want) in filter)
        {
            var k = key.ToLowerInvariant();
            switch (k)
            {
                case "author":
                    if (!string.Equals(rev.Author, want, StringComparison.Ordinal)) return false;
                    break;
                case "id":
                    if (!string.Equals(rev.Id, want, StringComparison.Ordinal)) return false;
                    break;
                case "type":
                    if (!RevisionTypeMatches(rev.Kind, want)) return false;
                    break;
                case "date":
                    // Exact date match — callers wanting ranges can filter
                    // outside Set via repeated calls. Date kept simple.
                    if (rev.Date == null) return false;
                    if (!DateTime.TryParse(want, out var w)) return false;
                    if (rev.Date.Value != w) return false;
                    break;
                default:
                    throw new ArgumentException(
                        $"unknown revision filter attribute `{key}` (valid: author, type, id, date)");
            }
        }
        return true;
    }

    /// <summary>Lenient type alias matching. `ins` matches insertion;
    /// `del` matches deletion; `format` matches formatChange and the
    /// structural *Change types when explicit; etc.</summary>
    private static bool RevisionTypeMatches(string actualKind, string want)
    {
        var w = want.ToLowerInvariant();
        var a = actualKind.ToLowerInvariant();
        if (a == w) return true;
        return (w, a) switch
        {
            ("ins", "insertion") => true,
            ("insertion", "insertion") => true,
            ("del", "deletion") => true,
            ("deletion", "deletion") => true,
            ("format", "formatchange") => true,
            ("format", "paragraphchange") => true,
            ("format", "sectionchange") => true,
            ("format", "tablechange") => true,
            ("format", "rowchange") => true,
            ("format", "cellchange") => true,
            ("paragraph", "paragraphchange") => true,
            ("rowins", "rowinsertion") => true,
            ("rowdel", "rowdeletion") => true,
            ("cellins", "cellinsertion") => true,
            ("celldel", "celldeletion") => true,
            ("paramarkins", "paragraphmarkinsertion") => true,
            ("paramarkdel", "paragraphmarkdeletion") => true,
            ("move", "movefrom") => true,
            ("move", "moveto") => true,
            _ => false,
        };
    }

    /// <summary>Accept one revision marker. Mirrors the per-kind branches in
    /// <see cref="AcceptAllChanges"/> but applied to a single element so the
    /// selector path can iterate matches with reverse-order safety.</summary>
    private void AcceptRevision(RevisionRef rev)
    {
        switch (rev.Kind)
        {
            case "insertion":
                {
                    var ins = (InsertedRun)rev.Element;
                    var parent = ins.Parent;
                    if (parent != null)
                    {
                        foreach (var child in ins.ChildElements.ToList())
                            parent.InsertBefore(child.CloneNode(true), ins);
                    }
                    ins.Remove();
                    break;
                }
            case "deletion":
                rev.Element.Remove();
                break;
            case "formatChange":
            case "paragraphChange":
            case "sectionChange":
            case "tableChange":
            case "rowChange":
            case "cellChange":
                rev.Element.Remove();
                break;
            case "rowInsertion":
            case "rowDeletion":
                // Accept row insertion: keep row, drop marker. Accept row
                // deletion: remove the entire row (the marker said "the row
                // was deleted", accept = commit the deletion).
                if (rev.Kind == "rowInsertion")
                {
                    rev.Element.Remove();
                }
                else
                {
                    var row = rev.Element.Ancestors<TableRow>().FirstOrDefault();
                    rev.Element.Remove();
                    row?.Remove();
                }
                break;
            case "cellInsertion":
                rev.Element.Remove();
                break;
            case "cellDeletion":
                {
                    var cell = rev.Element.Ancestors<TableCell>().FirstOrDefault();
                    rev.Element.Remove();
                    cell?.Remove();
                    break;
                }
            case "paragraphMarkInsertion":
                rev.Element.Remove();
                break;
            case "paragraphMarkDeletion":
                {
                    // Same merge logic as AcceptAllChanges: ¶ del means the
                    // paragraph break was deleted → join with next paragraph.
                    var pMark = (ParagraphMarkRunProperties)rev.Element.Parent!;
                    var thisPara = pMark.Ancestors<Paragraph>().FirstOrDefault();
                    var nextPara = thisPara?.NextSibling<Paragraph>();
                    rev.Element.Remove();
                    if (thisPara != null && nextPara != null)
                    {
                        var movable = thisPara.ChildElements
                            .Where(c => c is not ParagraphProperties)
                            .ToList();
                        var nextPPr = nextPara.GetFirstChild<ParagraphProperties>();
                        OpenXmlElement? insertAfter = nextPPr;
                        foreach (var ch in movable)
                        {
                            ch.Remove();
                            if (insertAfter == null) nextPara.PrependChild(ch);
                            else { insertAfter.InsertAfterSelf(ch); insertAfter = ch; }
                        }
                        thisPara.Remove();
                    }
                    break;
                }
            case "moveFrom":
                rev.Element.Remove();
                break;
            case "moveTo":
                {
                    var mt = (MoveToRun)rev.Element;
                    var parent = mt.Parent;
                    if (parent != null)
                    {
                        foreach (var child in mt.ChildElements.ToList())
                            parent.InsertBefore(child.CloneNode(true), mt);
                    }
                    mt.Remove();
                    break;
                }
        }
    }

    /// <summary>Reject one revision marker. Mirror of <see cref="AcceptRevision"/>
    /// with inverted semantics — discard inserts, restore deletes, restore
    /// prior pPr/rPr/etc. from each *Change's snapshot.</summary>
    private void RejectRevision(RevisionRef rev)
    {
        switch (rev.Kind)
        {
            case "insertion":
                rev.Element.Remove();
                break;
            case "deletion":
                {
                    var del = (DeletedRun)rev.Element;
                    var parent = del.Parent;
                    if (parent != null)
                    {
                        foreach (var child in del.ChildElements.ToList())
                        {
                            var clone = child.CloneNode(true);
                            foreach (var dt in clone.Descendants<DeletedText>().ToList())
                            {
                                var text = new Text(dt.Text);
                                if (dt.Space != null) text.Space = dt.Space;
                                dt.Parent?.ReplaceChild(text, dt);
                            }
                            parent.InsertBefore(clone, del);
                        }
                    }
                    del.Remove();
                    break;
                }
            case "formatChange":
                RestorePropsFromChange<RunProperties, RunPropertiesChange, PreviousRunProperties>(
                    (RunPropertiesChange)rev.Element, () => new RunProperties());
                break;
            case "paragraphChange":
                RestorePropsFromChange<ParagraphProperties, ParagraphPropertiesChange, ParagraphPropertiesExtended>(
                    (ParagraphPropertiesChange)rev.Element, () => new ParagraphProperties());
                break;
            case "sectionChange":
                RestorePropsFromChange<SectionProperties, SectionPropertiesChange, PreviousSectionProperties>(
                    (SectionPropertiesChange)rev.Element, () => new SectionProperties());
                break;
            case "tableChange":
                RestorePropsFromChange<TableProperties, TablePropertiesChange, PreviousTableProperties>(
                    (TablePropertiesChange)rev.Element, () => new TableProperties());
                break;
            case "cellChange":
                RestorePropsFromChange<TableCellProperties, TableCellPropertiesChange, PreviousTableCellProperties>(
                    (TableCellPropertiesChange)rev.Element, () => new TableCellProperties());
                break;
            case "rowChange":
                RestorePropsFromChange<TableRowProperties, TableRowPropertiesChange, PreviousTableRowProperties>(
                    (TableRowPropertiesChange)rev.Element, () => new TableRowProperties());
                break;
            case "rowInsertion":
                {
                    // Reject row insertion: discard the row entirely.
                    var row = rev.Element.Ancestors<TableRow>().FirstOrDefault();
                    rev.Element.Remove();
                    row?.Remove();
                    break;
                }
            case "rowDeletion":
                // Reject row deletion: keep row, drop marker.
                rev.Element.Remove();
                break;
            case "cellInsertion":
                {
                    var cell = rev.Element.Ancestors<TableCell>().FirstOrDefault();
                    rev.Element.Remove();
                    cell?.Remove();
                    break;
                }
            case "cellDeletion":
                rev.Element.Remove();
                break;
            case "paragraphMarkInsertion":
                {
                    // Reject ¶ ins: merge with previous paragraph.
                    var pMark = (ParagraphMarkRunProperties)rev.Element.Parent!;
                    rev.Element.Remove();
                    var thisPara = pMark.Ancestors<Paragraph>().FirstOrDefault();
                    var prevPara = thisPara?.PreviousSibling<Paragraph>();
                    if (thisPara != null && prevPara != null)
                    {
                        foreach (var ch in thisPara.ChildElements.Where(c => c is not ParagraphProperties).ToList())
                        {
                            ch.Remove();
                            prevPara.AppendChild(ch);
                        }
                        thisPara.Remove();
                    }
                    break;
                }
            case "paragraphMarkDeletion":
                rev.Element.Remove();
                break;
            case "moveFrom":
                {
                    var mf = (MoveFromRun)rev.Element;
                    var parent = mf.Parent;
                    if (parent != null)
                    {
                        foreach (var child in mf.ChildElements.ToList())
                            parent.InsertBefore(child.CloneNode(true), mf);
                    }
                    mf.Remove();
                    break;
                }
            case "moveTo":
                rev.Element.Remove();
                break;
        }
    }

    /// <summary>Pull the most representative human-readable text snippet for
    /// a revision marker. Mirrors what `query revision` shipped historically
    /// for ins/del/rPrChange/pPrChange and extends to the rest of the marker
    /// families. Empty when the marker has no associated text (e.g. structural
    /// section/table property changes that don't pin a specific run).</summary>
    private static string ExtractRevisionText(RevisionRef rev)
    {
        switch (rev.Kind)
        {
            case "insertion":
                return string.Join("", rev.Element.Descendants<Text>().Select(t => t.Text));
            case "deletion":
            case "moveFrom":
                return string.Join("", rev.Element.Descendants<DeletedText>().Select(t => t.Text));
            case "moveTo":
                return string.Join("", rev.Element.Descendants<Text>().Select(t => t.Text));
            case "formatChange":
                {
                    var run = rev.Element.Ancestors<Run>().FirstOrDefault();
                    return run != null
                        ? string.Join("", run.Descendants<Text>().Select(t => t.Text))
                        : "";
                }
            case "paragraphChange":
            case "paragraphMarkInsertion":
            case "paragraphMarkDeletion":
                {
                    var para = rev.Element.Ancestors<Paragraph>().FirstOrDefault();
                    return para != null
                        ? string.Join("", para.Descendants<Text>().Select(t => t.Text))
                        : "";
                }
            case "rowInsertion":
            case "rowDeletion":
            case "rowChange":
                {
                    var row = rev.Element.Ancestors<TableRow>().FirstOrDefault();
                    return row != null
                        ? string.Join(" | ",
                            row.Elements<TableCell>().Select(c =>
                                string.Join("", c.Descendants<Text>().Select(t => t.Text))))
                        : "";
                }
            case "cellInsertion":
            case "cellDeletion":
            case "cellChange":
                {
                    var cell = rev.Element.Ancestors<TableCell>().FirstOrDefault();
                    return cell != null
                        ? string.Join("", cell.Descendants<Text>().Select(t => t.Text))
                        : "";
                }
            case "tableChange":
            case "sectionChange":
                return "";
            default:
                return "";
        }
    }

    /// <summary>Compute the OOXML DOM path of the closest navigable ancestor
    /// of a revision marker. Useful for "where in the document is this
    /// revision" — agents can map a `/revision[N]` synthetic path back to a
    /// real `/body/p[@paraId=…]/r[K]` or `/body/tbl[N]/tr[M]/tc[K]` location
    /// for downstream `get` / cross-referencing.
    ///
    /// Path is best-effort and read-only: callers must NOT use it as a Set
    /// target to drive accept/reject (the legitimate path for that is
    /// `set /revision[N]` or a filtered selector — those go through the
    /// same EnumerateRevisions enumerator and are guaranteed to address
    /// the right marker even when revisions overlap on the same anchor).
    /// Empty when no navigable ancestor exists.</summary>
    private static string ComputeRevisionNativePath(OpenXmlElement marker)
    {
        // Walk up to find the closest "anchor" — a Run / Paragraph / TableCell
        // / TableRow / Table. The path is built from /body downward.
        var anchor = (OpenXmlElement?)marker;
        while (anchor != null
               && anchor is not Run
               && anchor is not Paragraph
               && anchor is not TableCell
               && anchor is not TableRow
               && anchor is not Table
               && anchor is not SectionProperties)
        {
            anchor = anchor.Parent;
        }
        if (anchor == null) return "";

        // Build segments root-down by walking back up.
        var segments = new List<string>();
        var current = anchor;
        while (current != null && current is not Body)
        {
            string? seg = current switch
            {
                Run r => $"r[{IndexOfSiblingsByType<Run>(r)}]",
                Paragraph p when !string.IsNullOrEmpty(p.ParagraphId?.Value)
                    => $"p[@paraId={p.ParagraphId!.Value}]",
                Paragraph p => $"p[{IndexOfSiblingsByType<Paragraph>(p)}]",
                TableCell tc => $"tc[{IndexOfSiblingsByType<TableCell>(tc)}]",
                TableRow tr => $"tr[{IndexOfSiblingsByType<TableRow>(tr)}]",
                Table tbl => $"tbl[{IndexOfSiblingsByType<Table>(tbl)}]",
                SectionProperties => "sectPr",
                _ => null,
            };
            if (seg != null) segments.Insert(0, seg);
            current = current.Parent;
        }
        if (segments.Count == 0) return "";
        return "/body/" + string.Join("/", segments);
    }

    private static int IndexOfSiblingsByType<T>(OpenXmlElement element) where T : OpenXmlElement
    {
        var parent = element.Parent;
        if (parent == null) return 1;
        int idx = 0;
        foreach (var sib in parent.Elements<T>())
        {
            idx++;
            if (ReferenceEquals(sib, element)) return idx;
        }
        return idx;
    }

    /// <summary>Common shape for <c>*PrChange</c> rejection: the change
    /// element's parent is the current pPr/rPr/etc.; the change's first child
    /// is the snapshot of the prior state. Replace the parent with a fresh
    /// element rebuilt from the snapshot's children.</summary>
    private static void RestorePropsFromChange<TParent, TChange, TSnapshot>(
        TChange change,
        Func<TParent> newInstance)
        where TParent : OpenXmlElement
        where TChange : OpenXmlElement
        where TSnapshot : OpenXmlElement
    {
        var parent = change.Parent as TParent;
        if (parent == null) { change.Remove(); return; }
        var snapshot = change.GetFirstChild<TSnapshot>();
        if (snapshot == null) { change.Remove(); return; }
        var grand = parent.Parent;
        if (grand == null) { change.Remove(); return; }
        var rebuilt = newInstance();
        foreach (var ch in snapshot.ChildElements.ToList())
            rebuilt.AppendChild(ch.CloneNode(true));
        grand.ReplaceChild(rebuilt, parent);
    }
}
