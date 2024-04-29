import string
from pprint import pprint

import fitz

if fitz.pymupdf_version_tuple < (1, 24, 0):
    raise NotImplementedError("PyMuPDF version 1.24.0 or later is needed.")


def to_markdown(doc: fitz.Document, pages: list = None) -> str:
    if isinstance(doc, str):
        doc = fitz.open(doc)
    SPACES = set(string.whitespace)
    if not pages:
        pages = range(doc.page_count)

    class IdentifyHeaders:

        def __init__(self, doc, pages: list = None, body_limit: float = None):
            if pages is None: 
                pages = range(doc.page_count)
            fontsizes = {}
            for pno in pages:
                page = doc[pno]
                blocks = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)["blocks"]
                for span in [ 
                    s
                    for b in blocks
                    for l in b["lines"]
                    for s in l["spans"]
                    if not SPACES.issuperset(s["text"])
                ]:
                    fontsz = round(span["size"])
                    count = fontsizes.get(fontsz, 0) + len(span["text"].strip())
                    fontsizes[fontsz] = count

            self.header_id = {}
            if body_limit is None: 
                body_limit = sorted(
                    [(k, v) for k, v in fontsizes.items()],
                    key=lambda i: i[1],
                    reverse=True,
                )[0][0]

            sizes = sorted(
                [f for f in fontsizes.keys() if f > body_limit], reverse=True
            )

            for i, size in enumerate(sizes):
                self.header_id[size] = "#" * (i + 1) + " "

        def get_header_id(self, span):
            fontsize = round(span["size"])
            hdr_id = self.header_id.get(fontsize, "")
            return hdr_id

    def resolve_links(links, span):
        bbox = fitz.Rect(span["bbox"])
        bbox_area = 0.7 * abs(bbox)
        for link in links:
            hot = link["from"]
            if not abs(hot & bbox) >= bbox_area:
                continue
            text = f'[{span["text"].strip()}]({link["uri"]})'
            return text

    def write_text(page, clip, hdr_prefix):
        out_string = ""
        code = False 
        links = [l for l in page.get_links() if l["kind"] == 2]

        blocks = page.get_text(
            "dict",
            clip=clip,
            flags=fitz.TEXTFLAGS_TEXT,
            sort=True,
        )["blocks"]

        for block in blocks:
            previous_y = 0
            for line in block["lines"]:
                if line["dir"][1] != 0:
                    continue
                spans = [s for s in line["spans"]]
                this_y = line["bbox"][3]
                same_line = abs(this_y - previous_y) <= 3 and previous_y > 0

                if same_line and out_string.endswith("\n"):
                    out_string = out_string[:-1]
                all_mono = all([s["flags"] & 8 for s in spans])
                text = "".join([s["text"] for s in spans])
                if not same_line:
                    previous_y = this_y
                    if not out_string.endswith("\n"):
                        out_string += "\n"

                if all_mono:
                    delta = int(
                        (spans[0]["bbox"][0] - block["bbox"][0])
                        / (spans[0]["size"] * 0.5)
                    )
                    if not code: 
                        out_string += "```"  
                        code = True
                    if not same_line:  
                        out_string += "\n" + " " * delta + text + " "
                        previous_y = this_y
                    else:  
                        out_string += text + " "
                    continue 

                for i, s in enumerate(spans):
                    if code: 
                        out_string += "```\n" 
                        code = False
                    mono = s["flags"] & 8
                    bold = s["flags"] & 16
                    italic = s["flags"] & 2

                    if mono:
                        out_string += f"`{s['text'].strip()}` "
                    else: 
                        if i == 0:
                            hdr_string = hdr_prefix.get_header_id(s)
                        else:
                            hdr_string = ""
                        prefix = ""
                        suffix = ""
                        if hdr_string == "":
                            if bold:
                                prefix = "**"
                                suffix += "**"
                            if italic:
                                prefix += "_"
                                suffix = "_" + suffix

                        ltext = resolve_links(links, s)
                        if ltext:
                            text = f"{hdr_string}{prefix}{ltext}{suffix} "
                        else:
                            text = f"{hdr_string}{prefix}{s['text'].strip()}{suffix} "
                        text = (
                            text.replace("<", "&lt;")
                            .replace(">", "&gt;")
                            .replace(chr(0xF0B7), "-")
                            .replace(chr(0xB7), "-")
                            .replace(chr(8226), "-")
                            .replace(chr(9679), "-")
                        )
                        out_string += text
                previous_y = this_y
                if not code:
                    out_string += "\n"
            out_string += "\n"
        if code:
            out_string += "```\n"
            code = False
        return out_string.replace(" \n", "\n")

    hdr_prefix = IdentifyHeaders(doc, pages=pages)
    md_string = ""

    for pno in pages:
        page = doc[pno]
        tabs = page.find_tables()
        tab_rects = sorted(
            [
                (fitz.Rect(t.bbox) | fitz.Rect(t.header.bbox), i)
                for i, t in enumerate(tabs.tables)
            ],
            key=lambda r: (r[0].y0, r[0].x0),
        )
        text_rects = []
        for i, (r, idx) in enumerate(tab_rects):
            if i == 0: 
                tr = page.rect
                tr.y1 = r.y0
                if not tr.is_empty:
                    text_rects.append(("text", tr, 0))
                text_rects.append(("table", r, idx))
                continue
            _, r0, idx0 = text_rects[-1]

            tr = page.rect
            tr.y0 = r0.y1
            tr.y1 = r.y0
            if not tr.is_empty:
                text_rects.append(("text", tr, 0))

            text_rects.append(("table", r, idx))

            if i == len(tab_rects) - 1:
                tr = page.rect
                tr.y0 = r.y1
                if not tr.is_empty:
                    text_rects.append(("text", tr, 0))

        if not text_rects:
            text_rects.append(("text", page.rect, 0))
        else:
            rtype, r, idx = text_rects[-1]
            if rtype == "table":
                tr = page.rect
                tr.y0 = r.y1
                if not tr.is_empty:
                    text_rects.append(("text", tr, 0))

        for rtype, r, idx in text_rects:
            if rtype == "text": 
                md_string += write_text(page, r, hdr_prefix) 
                md_string += "\n"
            else:  # a table rect
                md_string += tabs[idx].to_markdown(clean=False)

        md_string += "\n-----\n\n"

    return md_string


if __name__ == "__main__":
    import os
    import sys
    import time
    import pathlib

    try:
        filename = sys.argv[1]
    except IndexError:
        print(f"Usage:\npython {os.path.basename(__file__)} input.pdf")
        sys.exit()

    t0 = time.perf_counter()

    doc = fitz.open(filename)
    parms = sys.argv[2:]
    pages = range(doc.page_count)
    if len(parms) == 2 and parms[0] == "-pages":
        pages = []
        pages_spec = parms[1].replace("N", f"{doc.page_count}")
        for spec in pages_spec.split(","):
            if "-" in spec:
                start, end = map(int, spec.split("-"))
                pages.extend(range(start - 1, end))
            else:
                pages.append(int(spec) - 1)

        wrong_pages = set([n + 1 for n in pages if n >= doc.page_count][:4])
        if wrong_pages != set(): 
            sys.exit(f"Page number(s) {wrong_pages} not in '{doc}'.")

    md_string = to_markdown(doc, pages=pages)

    outname = doc.name.replace(".pdf", ".md")
    pathlib.Path(outname).write_bytes(md_string.encode())
    t1 = time.perf_counter()  # stop timer
    print(f"Markdown creation time for {doc.name=} {round(t1-t0,2)} sec.")
