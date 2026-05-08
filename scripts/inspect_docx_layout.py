import sys
from collections import Counter, defaultdict
from docx import Document


def pt(value):
    return None if value is None else round(value.pt, 2)


def describe_paragraph(paragraph):
    fmt = paragraph.paragraph_format
    return {
        "style": paragraph.style.name if paragraph.style else "None",
        "text": paragraph.text.strip()[:90],
        "before": pt(fmt.space_before),
        "after": pt(fmt.space_after),
        "left": pt(fmt.left_indent),
        "first": pt(fmt.first_line_indent),
        "align": str(fmt.alignment),
    }


def main(path):
    doc = Document(path)
    paragraphs = list(doc.paragraphs)
    style_stats = defaultdict(Counter)
    transitions = Counter()
    samples = []

    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue
        style = paragraph.style.name if paragraph.style else "None"
        if style in {"Heading 1", "Heading 2", "Heading 3", "Heading 4", "Normal", "List Paragraph"}:
            fmt = paragraph.paragraph_format
            style_stats[style][(pt(fmt.space_before), pt(fmt.space_after), pt(fmt.left_indent), pt(fmt.first_line_indent), str(fmt.alignment))] += 1

        if style.startswith("Heading"):
            next_non_empty = None
            for nxt in paragraphs[index + 1:]:
                if nxt.text.strip():
                    next_non_empty = nxt
                    break
            if next_non_empty is not None:
                next_style = next_non_empty.style.name if next_non_empty.style else "None"
                transitions[(style, next_style)] += 1
                if len(samples) < 60:
                    samples.append((describe_paragraph(paragraph), describe_paragraph(next_non_empty)))

    print("STYLE FORMAT STATS")
    for style in ["Heading 1", "Heading 2", "Heading 3", "Heading 4", "Normal", "List Paragraph"]:
        print(f"\n{style}")
        for key, count in style_stats[style].most_common(10):
            print(f"  {count:4d} before={key[0]} after={key[1]} left={key[2]} first={key[3]} align={key[4]}")

    print("\nHEADING -> NEXT CONTENT")
    for (style, next_style), count in transitions.most_common():
        print(f"  {count:4d} {style} -> {next_style}")

    print("\nSAMPLES")
    for heading, nxt in samples:
        print(f"H {heading['style']} b/a/l={heading['before']}/{heading['after']}/{heading['left']} :: {heading['text']}")
        print(f"N {nxt['style']} b/a/l={nxt['before']}/{nxt['after']}/{nxt['left']} :: {nxt['text']}")


if __name__ == "__main__":
    main(sys.argv[1])
