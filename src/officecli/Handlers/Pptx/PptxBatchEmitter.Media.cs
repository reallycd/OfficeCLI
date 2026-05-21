// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(picture-inline-base64): mirrors
    // WordBatchEmitter.Paragraph.TryEmitPictureRun — no size threshold, no
    // sidecar file, always emit `src="data:<contentType>;base64,<bytes>"`.
    // A 50MB picture produces a 70MB batch JSON; accepted by design.
    private static void EmitPicture(PowerPointHandler ppt, DocumentNode picNode,
                                    string parentSlidePath, string replayPath,
                                    List<BatchItem> items,
                                    SlideEmitContext ctx)
    {
        var fullPic = ppt.Get(picNode.Path);
        var props = FilterEmittableProps(fullPic.Format);
        DeferSlideJumpLink(props, replayPath, ctx);

        var binary = ppt.GetImageBinary(picNode.Path);
        if (binary.HasValue)
        {
            var (bytes, contentType) = binary.Value;
            props["src"] = $"data:{contentType};base64,{Convert.ToBase64String(bytes)}";
        }
        else
        {
            // No embedded part — picture is unresolvable on round-trip.
            // Drop to an unsupported warning rather than emit a half-row
            // that AddPicture would reject for missing src.
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: "picture",
                SlidePath: parentSlidePath,
                Reason: "picture has no resolvable embedded image part"));
            return;
        }

        // Drop Get-only diagnostic keys that AddPicture neither expects nor
        // accepts (mirrors docx WordBatchEmitter picture emit).
        props.Remove("id");
        props.Remove("contentType");
        props.Remove("fileSize");
        props.Remove("alt");
        // Re-add alt only if it was the explicit user-set value (not the
        // "(missing)" placeholder PictureToNode stamps in).
        var altRaw = fullPic.Format.TryGetValue("alt", out var av) ? av?.ToString() : null;
        if (!string.IsNullOrEmpty(altRaw) && altRaw != "(missing)")
            props["alt"] = altRaw;

        // Schema declares brightness/contrast/shadow/glow as add:false, set:true
        // on pptx/picture. AddPicture rejects them with UNSUPPORTED on replay
        // and the values are silently lost. Lift them out of the add bag and
        // defer to a follow-up `set` on the same replay path. Mirrors the
        // DeferSlideJumpLink pattern (deferred-set after every add).
        // Hard-coded picture-level drop list — same precedent as `image=true`,
        // `background=image`, `fill=gradient` drops elsewhere in this emitter.
        var deferredEffects = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in PictureSetOnlyEffectKeys)
        {
            if (props.TryGetValue(key, out var val))
            {
                deferredEffects[key] = val;
                props.Remove(key);
            }
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "picture",
            Props = props.Count > 0 ? props : null,
        });

        if (deferredEffects.Count > 0)
        {
            ctx.DeferredLinks.Add(new BatchItem
            {
                Command = "set",
                Path = replayPath,
                Props = deferredEffects,
            });
        }
    }

    // Picture effect props with schema `add: false, set: true`. Must NOT ride
    // along inside the add picture op props bag — AddPicture rejects them.
    private static readonly HashSet<string> PictureSetOnlyEffectKeys =
        new(StringComparer.OrdinalIgnoreCase)
    {
        "brightness", "contrast", "shadow", "glow",
    };
}
