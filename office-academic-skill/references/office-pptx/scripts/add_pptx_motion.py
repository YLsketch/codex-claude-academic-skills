#!/usr/bin/env python3
"""Add native PowerPoint transitions and entrance animations to a PPTX.

The script intentionally uses only the Python standard library. It opens a
PPTX as a zip package, updates ppt/slides/slide*.xml, and writes a new PPTX.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"

NS = {"p": P_NS, "a": A_NS, "r": R_NS, "p14": P14_NS}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


ANIMATION_FILTERS = {
    "appear": None,
    "fade": "fade",
    "wipe": "wipe(left)",
    "fly": "slide(fromBottom)",
    "zoom": "image",
}

PRESET_IDS = {
    "appear": (1, 0),
    "fade": (10, 0),
    "fly": (2, 4),
    "zoom": (23, 0),
    "wipe": (22, 1),
}

TRANSITIONS = {
    "cut": ("cut", {}),
    "fade": ("fade", {}),
    "push": ("push", {"dir": "r"}),
    "wipe": ("wipe", {"dir": "r"}),
    "split": ("split", {"orient": "horz", "dir": "out"}),
    "cover": ("cover", {"dir": "r"}),
    "dissolve": ("dissolve", {}),
    "circle": ("circle", {}),
    "diamond": ("diamond", {}),
    "plus": ("plus", {}),
    "wedge": ("wedge", {}),
}

CHROME_TOKENS = {
    "background",
    "bg",
    "header",
    "footer",
    "logo",
    "watermark",
    "pagenumber",
    "pagenum",
    "page",
    "date",
    "decor",
    "decoration",
    "decorative",
    "rule",
}


@dataclass(frozen=True)
class Target:
    shape_id: int
    name: str
    order: int


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def normalize_name(value: str) -> list[str]:
    compact = re.sub(r"[^a-z0-9]+", " ", value.lower())
    return [token for token in compact.split() if token]


def is_chrome(name: str) -> bool:
    tokens = normalize_name(name)
    return any(token in CHROME_TOKENS for token in tokens)


def parse_slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def copy_zip_to_dir(pptx_path: Path, work_dir: Path) -> None:
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(work_dir)


def write_dir_to_zip(source_dir: Path, output_path: Path) -> None:
    if output_path.exists():
        output_path.unlink()
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(source_dir):
            for filename in files:
                file_path = Path(root) / filename
                arcname = file_path.relative_to(source_dir).as_posix()
                zf.write(file_path, arcname)


def find_slide_paths(work_dir: Path) -> list[Path]:
    slides_dir = work_dir / "ppt" / "slides"
    if not slides_dir.is_dir():
        raise FileNotFoundError("ppt/slides directory not found; input is not a normal PPTX")
    presentation = ET.parse(work_dir / "ppt" / "presentation.xml").getroot()
    rels_root = ET.parse(work_dir / "ppt" / "_rels" / "presentation.xml.rels").getroot()
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rel_map = {rel.get("Id"): rel.get("Target") for rel in rels_root.findall(f"{{{rel_ns}}}Relationship")}
    slide_list = presentation.find("./p:sldIdLst", NS)
    ordered = []
    if slide_list is not None:
        for slide_id in slide_list.findall("./p:sldId", NS):
            rel_id = slide_id.get(f"{{{R_NS}}}id")
            target = rel_map.get(rel_id)
            if not target:
                continue
            if target.startswith("/"):
                slide_path = work_dir / target.lstrip("/")
            else:
                slide_path = work_dir / "ppt" / target
            if slide_path.is_file():
                ordered.append(slide_path)
    return ordered or sorted(slides_dir.glob("slide*.xml"), key=lambda p: parse_slide_number(p.name))


def get_slide_size(work_dir: Path) -> tuple[int, int]:
    presentation = ET.parse(work_dir / "ppt" / "presentation.xml").getroot()
    size = presentation.find("./p:sldSz", NS)
    if size is None:
        return (0, 0)
    return int(size.get("cx", "0")), int(size.get("cy", "0"))


def get_cnvpr_id_and_name(elem: ET.Element) -> tuple[int | None, str]:
    cnvpr = elem.find(".//p:cNvPr", NS)
    if cnvpr is None:
        return None, ""
    raw_id = cnvpr.get("id")
    if raw_id is None:
        return None, cnvpr.get("name", "")
    try:
        return int(raw_id), cnvpr.get("name", "")
    except ValueError:
        return None, cnvpr.get("name", "")


def is_flattened_full_slide_picture(root: ET.Element, slide_size: tuple[int, int]) -> bool:
    sp_tree = root.find("./p:cSld/p:spTree", NS)
    if sp_tree is None or not all(slide_size):
        return False
    allowed = {qn("p", tag) for tag in ("sp", "pic", "graphicFrame", "grpSp")}
    objects = [child for child in list(sp_tree) if child.tag in allowed]
    if len(objects) != 1 or objects[0].tag != qn("p", "pic"):
        return False
    xfrm = objects[0].find("./p:spPr/a:xfrm", NS)
    if xfrm is None:
        return False
    off = xfrm.find("./a:off", NS)
    ext = xfrm.find("./a:ext", NS)
    if off is None or ext is None:
        return False
    x, y = int(off.get("x", "0")), int(off.get("y", "0"))
    cx, cy = int(ext.get("cx", "0")), int(ext.get("cy", "0"))
    sw, sh = slide_size
    tolerance = max(36576, int(min(sw, sh) * 0.015))
    return abs(x) <= tolerance and abs(y) <= tolerance and abs(cx - sw) <= tolerance and abs(cy - sh) <= tolerance


def discover_targets(root: ET.Element, include_chrome: bool, max_targets: int) -> list[Target]:
    sp_tree = root.find("./p:cSld/p:spTree", NS)
    if sp_tree is None:
        return []

    targets: list[Target] = []
    allowed = {qn("p", "sp"), qn("p", "pic"), qn("p", "graphicFrame"), qn("p", "grpSp")}
    for child in list(sp_tree):
        if child.tag not in allowed:
            continue
        shape_id, name = get_cnvpr_id_and_name(child)
        if shape_id is None:
            continue
        if not include_chrome and is_chrome(name):
            continue
        targets.append(Target(shape_id=shape_id, name=name, order=len(targets) + 1))
        if max_targets and len(targets) >= max_targets:
            break
    return targets


def remove_existing(root: ET.Element) -> None:
    for child in list(root):
        if child.tag in {qn("p", "transition"), qn("p", "timing")}:
            root.remove(child)


def insertion_index(root: ET.Element) -> int:
    ext_index = None
    last_after_core = -1
    for idx, child in enumerate(list(root)):
        if child.tag == qn("p", "extLst"):
            ext_index = idx
            break
        if child.tag in {qn("p", "cSld"), qn("p", "clrMapOvr")}:
            last_after_core = idx
    return ext_index if ext_index is not None else last_after_core + 1


def build_transition(effect: str, duration: float, auto_advance: float | None) -> ET.Element:
    elem_name, attrs = TRANSITIONS[effect]
    transition = ET.Element(qn("p", "transition"))
    transition.set(qn("p14", "dur"), str(int(duration * 1000)))
    if auto_advance is not None:
        transition.set("advTm", str(int(auto_advance * 1000)))
    ET.SubElement(transition, qn("p", elem_name), attrs)
    return transition


def ctn(elem_id: int, **attrs: str) -> ET.Element:
    attrs = {"id": str(elem_id), **attrs}
    return ET.Element(qn("p", "cTn"), attrs)


def add_start_condition(parent: ET.Element, delay: str) -> None:
    st = ET.SubElement(parent, qn("p", "stCondLst"))
    ET.SubElement(st, qn("p", "cond"), {"delay": delay})


def build_effect_block(shape_id: int, animation: str, duration_ms: int, next_id: int) -> tuple[list[ET.Element], int]:
    set_id = next_id
    eff_id = next_id + 1
    next_id += 2

    set_elem = ET.Element(qn("p", "set"))
    c_bhvr = ET.SubElement(set_elem, qn("p", "cBhvr"))
    set_ctn = ET.SubElement(c_bhvr, qn("p", "cTn"), {"id": str(set_id), "dur": "1", "fill": "hold"})
    add_start_condition(set_ctn, "0")
    tgt = ET.SubElement(c_bhvr, qn("p", "tgtEl"))
    ET.SubElement(tgt, qn("p", "spTgt"), {"spid": str(shape_id)})
    attr_lst = ET.SubElement(c_bhvr, qn("p", "attrNameLst"))
    ET.SubElement(attr_lst, qn("p", "attrName")).text = "style.visibility"
    to = ET.SubElement(set_elem, qn("p", "to"))
    ET.SubElement(to, qn("p", "strVal"), {"val": "visible"})

    filter_name = ANIMATION_FILTERS[animation]
    if filter_name is None:
        return [set_elem], next_id

    anim = ET.Element(qn("p", "animEffect"), {"transition": "in", "filter": filter_name})
    c_bhvr = ET.SubElement(anim, qn("p", "cBhvr"))
    ET.SubElement(c_bhvr, qn("p", "cTn"), {"id": str(eff_id), "dur": str(duration_ms)})
    tgt = ET.SubElement(c_bhvr, qn("p", "tgtEl"))
    ET.SubElement(tgt, qn("p", "spTgt"), {"spid": str(shape_id)})
    return [set_elem, anim], next_id


def grouped_target_indexes(target_count: int, *, clicks_per_slide: int, group_size: int) -> set[int]:
    if target_count <= 0:
        return set()
    if group_size > 0:
        return {idx for idx in range(target_count) if idx % group_size == 0}
    clicks = max(1, clicks_per_slide)
    inferred_group_size = max(1, (target_count + clicks - 1) // clicks)
    return {idx for idx in range(target_count) if idx % inferred_group_size == 0}


def build_timing(
    targets: list[Target],
    animation: str,
    duration: float,
    stagger: float,
    trigger: str,
    clicks_per_slide: int,
    group_size: int,
) -> ET.Element:
    duration_ms = int(duration * 1000)
    stagger_ms = int(stagger * 1000)
    next_id = 3

    timing = ET.Element(qn("p", "timing"))
    tn_lst = ET.SubElement(timing, qn("p", "tnLst"))
    root_par = ET.SubElement(tn_lst, qn("p", "par"))
    root_ctn = ET.SubElement(root_par, qn("p", "cTn"), {
        "id": "1",
        "dur": "indefinite",
        "restart": "never",
        "nodeType": "tmRoot",
    })
    child_lst = ET.SubElement(root_ctn, qn("p", "childTnLst"))
    seq = ET.SubElement(child_lst, qn("p", "seq"), {"concurrent": "1", "nextAc": "seek"})
    main = ET.SubElement(seq, qn("p", "cTn"), {"id": "2", "dur": "indefinite", "nodeType": "mainSeq"})
    main_children = ET.SubElement(main, qn("p", "childTnLst"))

    elapsed_ms = 0
    group_starts = grouped_target_indexes(
        len(targets),
        clicks_per_slide=clicks_per_slide,
        group_size=group_size,
    )
    for index, target in enumerate(targets):
        preset_id, preset_subtype = PRESET_IDS[animation]
        wrapper_par = ET.SubElement(main_children, qn("p", "par"))
        wrapper = ctn(next_id, fill="hold")
        next_id += 1
        wrapper_par.append(wrapper)

        if trigger == "on-click":
            add_start_condition(wrapper, "indefinite")
            node_type = "clickEffect"
        elif trigger == "grouped-clicks":
            if index in group_starts:
                elapsed_ms = 0
                add_start_condition(wrapper, "indefinite")
                node_type = "clickEffect"
            else:
                elapsed_ms += duration_ms + stagger_ms
                add_start_condition(wrapper, str(elapsed_ms))
                node_type = "afterEffect"
        elif trigger == "with-previous":
            add_start_condition(wrapper, "0")
            node_type = "withEffect"
        else:
            if index == 0:
                elapsed_ms = 0
            else:
                elapsed_ms += duration_ms + stagger_ms
            add_start_condition(wrapper, str(elapsed_ms))
            node_type = "afterEffect"

        wrapper_children = ET.SubElement(wrapper, qn("p", "childTnLst"))
        leaf_par = ET.SubElement(wrapper_children, qn("p", "par"))
        leaf = ET.SubElement(leaf_par, qn("p", "cTn"), {
            "id": str(next_id),
            "presetID": str(preset_id),
            "presetClass": "entr",
            "presetSubtype": str(preset_subtype),
            "fill": "hold",
            "nodeType": node_type,
        })
        next_id += 1
        add_start_condition(leaf, "0")
        leaf_children = ET.SubElement(leaf, qn("p", "childTnLst"))
        blocks, next_id = build_effect_block(target.shape_id, animation, duration_ms, next_id)
        for block in blocks:
            leaf_children.append(block)

    prev = ET.SubElement(seq, qn("p", "prevCondLst"))
    cond = ET.SubElement(prev, qn("p", "cond"), {"evt": "onPrev", "delay": "0"})
    tgt = ET.SubElement(cond, qn("p", "tgtEl"))
    ET.SubElement(tgt, qn("p", "sldTgt"))

    nxt = ET.SubElement(seq, qn("p", "nextCondLst"))
    cond = ET.SubElement(nxt, qn("p", "cond"), {"evt": "onNext", "delay": "0"})
    tgt = ET.SubElement(cond, qn("p", "tgtEl"))
    ET.SubElement(tgt, qn("p", "sldTgt"))

    bld_lst = ET.SubElement(timing, qn("p", "bldLst"))
    for target in targets:
        ET.SubElement(bld_lst, qn("p", "bldP"), {"spid": str(target.shape_id), "grpId": "0"})
    return timing


def process_slide(
    slide_path: Path,
    *,
    slide_number: int,
    slide_size: tuple[int, int],
    transition: str,
    transition_duration: float,
    auto_advance: float | None,
    animation: str,
    duration: float,
    stagger: float,
    trigger: str,
    include_chrome: bool,
    max_targets: int,
    replace_existing: bool,
    allow_full_slide_image: bool,
    clicks_per_slide: int,
    group_size: int,
) -> tuple[int, list[Target]]:
    tree = ET.parse(slide_path)
    root = tree.getroot()
    targets = discover_targets(root, include_chrome=include_chrome, max_targets=max_targets)
    had_transition = root.find("./p:transition", NS) is not None
    had_timing = root.find("./p:timing", NS) is not None
    flattened = is_flattened_full_slide_picture(root, slide_size)

    if replace_existing:
        remove_existing(root)
        had_transition = had_timing = False

    idx = insertion_index(root)
    if transition != "none" and not had_transition:
        root.insert(idx, build_transition(transition, transition_duration, auto_advance))
        idx += 1
    if animation != "none" and targets and not had_timing and (not flattened or allow_full_slide_image):
        root.insert(idx, build_timing(targets, animation, duration, stagger, trigger, clicks_per_slide, group_size))

    if had_transition and not replace_existing:
        print(f"slide{slide_number}: existing transition preserved")
    if had_timing and not replace_existing:
        print(f"slide{slide_number}: existing element timing preserved")
    if flattened and animation != "none" and not allow_full_slide_image:
        print(f"slide{slide_number}: flattened full-slide picture detected; element animation skipped (transition can still be added)")

    tree.write(slide_path, encoding="UTF-8", xml_declaration=True)
    return slide_number, targets


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Inject native PowerPoint animations into a PPTX.")
    parser.add_argument("input", help="Source .pptx file")
    parser.add_argument("-o", "--output", help="Output .pptx file. Defaults to <input>-animated.pptx")
    parser.add_argument("--transition", choices=["none", *TRANSITIONS.keys()], default="fade")
    parser.add_argument("--transition-duration", type=float, default=0.4)
    parser.add_argument("--auto-advance", type=float, default=None, help="Auto advance slides after N seconds")
    parser.add_argument("--effect", choices=["none", *ANIMATION_FILTERS.keys()], default="fade")
    parser.add_argument("--duration", type=float, default=0.45, help="Object animation duration in seconds")
    parser.add_argument("--stagger", type=float, default=0.12, help="Gap between objects inside a click group")
    parser.add_argument(
        "--trigger",
        choices=["grouped-clicks", "after-previous", "with-previous", "on-click"],
        default="grouped-clicks",
        help="Object reveal behavior. grouped-clicks creates 1-2 presenter clicks per slide by default.",
    )
    parser.add_argument("--clicks-per-slide", type=int, default=2, help="Target click groups per slide for --trigger grouped-clicks")
    parser.add_argument("--group-size", type=int, default=0, help="Objects per click group for --trigger grouped-clicks; 0 infers from --clicks-per-slide")
    parser.add_argument("--include-chrome", action="store_true", help="Animate likely headers, footers, logos, and decorations")
    parser.add_argument("--max-targets", type=int, default=12, help="Maximum animated objects per slide; 0 means no cap")
    parser.add_argument("--replace-existing", action="store_true", help="Replace existing transition/timing XML; default is to preserve it")
    parser.add_argument("--allow-full-slide-image", action="store_true", help="Allow whole-slide image entrance; this is not per-element animation")
    parser.add_argument("--skip-element-slides", default="", help="Comma-separated 1-based slide numbers to skip for element animation, e.g. 1,2,12")
    parser.add_argument("--overwrite-output", action="store_true", help="Allow replacing an existing output file")
    parser.add_argument("--dry-run", action="store_true", help="Print target summary without writing output")
    return parser


def validate_args(args: argparse.Namespace) -> tuple[Path, Path]:
    source = Path(args.input).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"Input PPTX not found: {source}")
    if source.suffix.lower() != ".pptx":
        raise ValueError("Input must be a .pptx file")
    if args.transition_duration < 0 or args.duration < 0 or args.stagger < 0:
        raise ValueError("Durations and stagger must be non-negative")
    if args.auto_advance is not None and args.auto_advance < 0:
        raise ValueError("--auto-advance must be non-negative")
    if args.clicks_per_slide < 1:
        raise ValueError("--clicks-per-slide must be at least 1")
    if args.group_size < 0:
        raise ValueError("--group-size must be non-negative")
    if args.output:
        output = Path(args.output).expanduser().resolve()
    else:
        output = source.with_name(f"{source.stem}-animated{source.suffix}")
    if output == source and not args.dry_run:
        raise ValueError("Output must differ from input; refusing to overwrite source")
    if output.exists() and not args.overwrite_output and not args.dry_run:
        raise FileExistsError(f"Output already exists: {output}; choose another name or pass --overwrite-output")
    return source, output


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        source, output = validate_args(args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    with tempfile.TemporaryDirectory(prefix="pptx-animations-") as temp:
        work_dir = Path(temp) / "pptx"
        copy_zip_to_dir(source, work_dir)
        slide_paths = find_slide_paths(work_dir)
        slide_size = get_slide_size(work_dir)
        try:
            skip_element_slides = {int(item.strip()) for item in args.skip_element_slides.split(",") if item.strip()}
        except ValueError as exc:
            raise ValueError("--skip-element-slides must be a comma-separated list of slide numbers") from exc
        if any(number < 1 or number > len(slide_paths) for number in skip_element_slides):
            raise ValueError("--skip-element-slides contains a slide number outside the presentation")
        summaries: list[tuple[int, list[Target]]] = []
        for slide_number, slide_path in enumerate(slide_paths, start=1):
            slide_effect = "none" if slide_number in skip_element_slides else args.effect
            summaries.append(
                process_slide(
                    slide_path,
                    slide_number=slide_number,
                    slide_size=slide_size,
                    transition=args.transition,
                    transition_duration=args.transition_duration,
                    auto_advance=args.auto_advance,
                    animation=slide_effect,
                    duration=args.duration,
                    stagger=args.stagger,
                    trigger=args.trigger,
                    include_chrome=args.include_chrome,
                    max_targets=args.max_targets,
                    replace_existing=args.replace_existing,
                    allow_full_slide_image=args.allow_full_slide_image,
                    clicks_per_slide=args.clicks_per_slide,
                    group_size=args.group_size,
                )
            )

        for slide_num, targets in summaries:
            names = ", ".join(f"{target.shape_id}:{target.name or 'unnamed'}" for target in targets)
            print(f"slide{slide_num}: {len(targets)} target(s)" + (f" [{names}]" if names else ""))

        if args.dry_run:
            print("Dry run complete; no output written.")
            return 0

        output.parent.mkdir(parents=True, exist_ok=True)
        write_dir_to_zip(work_dir, output)
        if not output.exists():
            raise RuntimeError(f"Output was not created: {output}")
        print(f"Animated PPTX written: {output}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
