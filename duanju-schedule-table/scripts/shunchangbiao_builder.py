#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import math
import re
import shutil
import subprocess
import tempfile
from collections import OrderedDict
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
except ImportError:  # pragma: no cover - dependency depends on the active runtime
    pdfplumber = None


CN_DIGITS = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
CN_UNITS = {"十": 10, "百": 100, "千": 1000}
EPISODE_RE = r"^\s*第([\d零〇一二两三四五六七八九十百千]+)集"

GROUP_HINTS = (
    "若干",
    "众人",
    "员工",
    "职员",
    "应聘者",
    "高层",
    "保镖",
    "村民",
    "群众",
    "游客",
    "干部",
    "工作人员",
    "代表",
    "老村民",
    "股东",
    "参会人",
    "客人",
    "夜枭成员",
    "保安",
    "队员",
    "舞女",
    "劳工",
    "流民",
    "家人",
)

DAY_RE = r"日|夜|清晨|上午|下午|傍晚|深夜|凌晨"
IO_RE = r"内外|内|外"

PROP_KEYWORDS = [
    ("轮椅", "轮椅"),
    ("沈家纹路金属残片", "沈家纹路金属残片"),
    ("金属残片", "金属残片"),
    ("全息光屏", "全息光屏画面"),
    ("光屏", "光屏画面"),
    ("卷宗", "卷宗"),
    ("口红", "口红"),
    ("酒杯", "酒杯"),
    ("长刀", "长刀"),
    ("机械义肢", "机械义肢"),
    ("井盖", "井盖"),
    ("手电筒", "手电筒"),
    ("防爆盾", "防爆盾"),
    ("电棍", "电棍"),
    ("子弹", "子弹/枪械效果"),
    ("狙击手", "枪械效果"),
    ("婴儿车", "婴儿车"),
    ("奶瓶", "奶瓶"),
    ("纸尿裤", "纸尿裤"),
    ("襁褓", "襁褓"),
    ("书信", "书信"),
    ("推荐信", "推荐信"),
    ("门把手", "门把手"),
    ("清单", "清单"),
    ("运动衣", "运动衣"),
    ("菜刀", "菜刀"),
    ("鞭子", "鞭子"),
    ("遥控器", "遥控器"),
    ("炸弹", "炸弹道具"),
    ("棍棒", "棍棒"),
    ("竞标方案", "竞标方案"),
    ("投影", "投影屏"),
    ("公章", "公章"),
    ("股权书", "股权书"),
    ("短刀", "短刀"),
    ("银针", "银针"),
    ("横幅", "横幅"),
    ("手机", "手机"),
    ("电话", "手机/电话"),
    ("平板", "平板"),
    ("电脑", "电脑"),
    ("文件夹", "文件夹"),
    ("文件", "文件"),
    ("合同", "合同"),
    ("协议", "协议"),
    ("信封", "信封"),
    ("现金", "现金"),
    ("两千块", "现金"),
    ("验孕棒", "验孕棒"),
    ("戒指", "戒指"),
    ("首饰盒", "首饰盒"),
    ("蛋糕", "蛋糕"),
    ("菜肴", "菜肴"),
    ("饮料", "饮料"),
    ("项链", "项链"),
    ("购物袋", "购物袋"),
    ("保温壶", "保温壶"),
    ("咖啡杯", "咖啡杯"),
    ("咖啡", "咖啡"),
    ("粉包", "粉包"),
    ("照片", "照片"),
    ("报告", "报告"),
    ("亲子鉴定", "亲子鉴定报告"),
    ("拐杖", "拐杖"),
    ("对讲机", "对讲机"),
    ("椅子", "椅子"),
    ("玻璃瓶", "玻璃瓶"),
    ("豪车", "豪车"),
    ("清洁工具", "清洁工具"),
    ("水桶", "水桶"),
    ("冷水", "冷水"),
    ("棉签", "棉签"),
    ("药品", "药品"),
    ("擦点药", "药品"),
    ("病床", "病床"),
    ("输液", "输液道具"),
    ("水杯", "水杯"),
    ("糕点", "糕点"),
    ("水果刀", "水果刀"),
    ("美工刀", "美工刀"),
    ("刀", "刀具"),
    ("花瓶", "花瓶"),
    ("财务报表", "财务报表"),
    ("催债函", "催债函"),
]


def cn_episode_to_int(value: str) -> int | None:
    if value.isdigit():
        return int(value)
    if not value or any(char not in CN_DIGITS and char not in CN_UNITS for char in value):
        return None
    total = 0
    current = 0
    for char in value:
        if char in CN_DIGITS:
            current = CN_DIGITS[char]
            continue
        unit = CN_UNITS[char]
        total += (current or 1) * unit
        current = 0
    return total + current


def normalize_scene_heading(text: str) -> str:
    text = text.replace("　", " ").strip()
    text = re.sub(r"^[•·●○▪■□◆◇◦\-\s]+", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def clean_location(location: str) -> str:
    location = location.strip(" 　-=—")
    location = re.sub(r"\s*-\s*", "-", location)
    location = re.sub(r"\s+", " ", location)
    return location


def is_scene_heading(text: str) -> bool:
    text = normalize_scene_heading(text)
    if not text:
        return False
    if re.match(r"^\d+\s*[-—–－]\s*\d+", text):
        return True
    if any(mark in text for mark in "：:。？！；;"):
        return False
    generated_day_first = rf"^(?:{DAY_RE})\s*(?:{IO_RE})\s*.+$"
    generated_location_first = rf"^.+?(?:{DAY_RE})\s*(?:{IO_RE})$"
    return bool(re.match(generated_day_first, text) or re.match(generated_location_first, text))


def parse_heading(
    text: str,
    episode: int | None,
    scene_count: dict[int, int],
    last_day: str,
    last_inside_outside: str,
) -> tuple[str, str, str, str, str, str] | None:
    text = normalize_scene_heading(text)
    explicit = re.match(r"^(\d+)\s*[-—–－]\s*(\d+)\s*(.*)$", text)
    if explicit:
        ep = int(explicit.group(1))
        no = int(explicit.group(2))
        scene_count[ep] = max(scene_count.get(ep, 0), no)
        remainder = explicit.group(3).strip()
        heading_note = ""
        note_match = re.search(r"\s*[（(]([^）)]*)[）)]\s*$", remainder)
        if note_match:
            heading_note = note_match.group(1).strip()
            remainder = remainder[: note_match.start()].strip()

        day = ""
        inside_outside = ""
        location = remainder
        day_first = re.match(rf"^(?P<day>{DAY_RE})\s*(?P<io>{IO_RE})\s*(?P<location>.+)$", remainder)
        location_first = re.match(rf"^(?P<location>.+?)\s*(?P<day>{DAY_RE})\s*(?P<io>{IO_RE})$", remainder)
        io_first = re.match(rf"^(?P<io>{IO_RE})\s*(?P<location>.+)$", remainder)
        if day_first:
            day = normalize_day(day_first.group("day"))
            inside_outside = day_first.group("io")
            location = day_first.group("location")
        elif location_first:
            day = normalize_day(location_first.group("day"))
            inside_outside = location_first.group("io")
            location = location_first.group("location")
        elif io_first:
            inside_outside = io_first.group("io")
            location = io_first.group("location")

        warnings = []
        if not day:
            if last_day in {"日", "夜"}:
                day = last_day
                warnings.append("日夜按上一场推断")
            else:
                day = "需确认"
                warnings.append("日夜需确认")
        if not inside_outside:
            if last_inside_outside in {"内", "外", "内外"}:
                inside_outside = last_inside_outside
                warnings.append("内外按上一场推断")
            else:
                inside_outside = "需确认"
                warnings.append("内外需确认")
        if not location:
            location = "场景需确认"
            warnings.append("场景需确认")
        return f"{ep}-{no}", day, inside_outside, clean_location(location), "；".join(warnings), heading_note

    generated = re.match(rf"^({DAY_RE})\s*({IO_RE})\s*(.+)$", text)
    if generated and episode:
        scene_count[episode] = scene_count.get(episode, 0) + 1
        return f"{episode}-{scene_count[episode]}", normalize_day(generated.group(1)), generated.group(2), clean_location(generated.group(3)), "", ""

    generated_location_first = re.match(
        rf"^(?P<location>.+?)\s*(?P<day>{DAY_RE})\s*(?P<io>{IO_RE})$",
        text,
    )
    if generated_location_first and episode:
        scene_count[episode] = scene_count.get(episode, 0) + 1
        return (
            f"{episode}-{scene_count[episode]}",
            normalize_day(generated_location_first.group("day")),
            generated_location_first.group("io"),
            clean_location(generated_location_first.group("location")),
            "",
            "",
        )
    return None


def normalize_day(day: str) -> str:
    if day in {"清晨", "上午", "下午", "傍晚"}:
        return "日"
    if day in {"深夜", "凌晨"}:
        return "夜"
    return day


def normalize_person(name: str) -> str:
    name = re.sub(r"【[^】]*】", "", name)
    name = re.sub(r"[（(][^）)]*[）)]", "", name)
    name = re.sub(r"[*＊]\s*\d+$", "", name)
    name = name.strip(" 　、，,;；:：")
    if re.fullmatch(r"[Nn\d]+", name):
        return ""
    if name in {"人物", "人物小传", "人物小传："}:
        return ""
    return name


def split_people(line: str) -> list[str]:
    line = re.sub(r"^人物\s*[:：;；]?", "", line.strip())
    line = re.sub(r"([一-龥A-Za-z]+)\s+([Nn])(?=$|[、，,\s])", r"\1\2", line)
    line = re.sub(r"([一-龥A-Za-z]+)\s+(\d+)(?=$|[、，,\s])", r"\1", line)
    parts = re.split(r"[、，,\s]+", line)
    result = []
    for part in parts:
        name = normalize_person(part)
        if name:
            result.append(name)
    return result


def paragraph_text(paragraph: Paragraph) -> str:
    xml_text = "".join(paragraph._p.xpath(".//w:t/text()"))
    return (xml_text or paragraph.text).strip()


def docx_lines(path: Path) -> list[str]:
    document = Document(path)
    lines = []
    for block in document.iter_inner_content():
        if isinstance(block, Paragraph):
            text = paragraph_text(block)
            if text:
                lines.append(text)
            continue
        if isinstance(block, Table):
            seen_cells = set()
            for row in block.rows:
                for cell in row.cells:
                    cell_key = cell._tc
                    if cell_key in seen_cells:
                        continue
                    seen_cells.add(cell_key)
                    for paragraph in cell.paragraphs:
                        text = paragraph_text(paragraph)
                        if text:
                            lines.append(text)
    return lines


def augment_visible_people(scenes: list[dict]) -> None:
    candidates = ordered_unique(
        person
        for scene in scenes
        for person in scene.get("people", [])
        if not is_group_name(person) and not is_voice_name(person)
    )
    for scene in scenes:
        people = list(scene.get("people", []))
        for role in candidates:
            if role in people:
                continue
            if any(age_variant_base(person) == role for person in people):
                continue
            pattern = re.escape(role)
            for line in scene.get("body", []):
                text = line.lstrip("△").strip()
                if re.match(rf"^{pattern}VO\b", text, flags=re.IGNORECASE):
                    continue
                if re.match(rf"^{pattern}(?:[（(][^）)]*[）)])?[:：]", text) or (
                    line.startswith("△") and text.startswith(role)
                ):
                    people.append(role)
                    break
        variant_bases = {base for person in people if (base := age_variant_base(person))}
        scene_text = "\n".join(scene.get("body", []))
        if "闪回" in scene.get("heading_note", "") or any(
            re.search(rf"(?:幼年|年幼|少年|青年|老年){re.escape(base)}", scene_text)
            for base in variant_bases
        ):
            people = [person for person in people if person not in variant_bases]
        scene["people"] = people


def age_variant_base(name: str) -> str:
    match = re.match(r"^(?:幼年|年幼|少年|青年|老年)(.+)$", name)
    return match.group(1) if match else ""


def parse_docx(path: Path) -> list[dict]:
    paragraphs = docx_lines(path)
    has_body_marker = any(p.startswith("正文") for p in paragraphs)
    in_body = not has_body_marker
    episode = None
    scene_count: dict[int, int] = {}
    scenes: list[dict] = []
    last_day = ""
    last_inside_outside = ""

    for idx, text in enumerate(paragraphs):
        if text.startswith("正文"):
            in_body = True
            continue
        if not in_body:
            continue
        ep_match = re.match(EPISODE_RE, text)
        if ep_match:
            episode = cn_episode_to_int(ep_match.group(1))
            if episode is not None:
                scene_count.setdefault(episode, 0)
            continue
        if not is_scene_heading(text):
            continue

        parsed = parse_heading(text, episode, scene_count, last_day, last_inside_outside)
        if not parsed:
            continue
        scene_id, day, inside_outside, location, warning, heading_note = parsed
        if "-" in scene_id and scene_id.split("-", 1)[0].isdigit():
            episode = int(scene_id.split("-", 1)[0])
        if day in {"日", "夜"}:
            last_day = day
        if inside_outside in {"内", "外", "内外"}:
            last_inside_outside = inside_outside

        people_line = ""
        for lookahead in range(idx + 1, min(len(paragraphs), idx + 8)):
            target = paragraphs[lookahead]
            if target.startswith("人物"):
                people_line = target
                break
            if is_scene_heading(target) or re.match(EPISODE_RE, target):
                break

        scenes.append(
            {
                "source": path.name,
                "idx": idx,
                "heading": text,
                "scene": scene_id,
                "day": day,
                "inside_outside": inside_outside,
                "location": location,
                "people_raw": people_line,
                "people": split_people(people_line),
                "warning": warning,
                "heading_note": heading_note,
            }
        )

    for i, scene in enumerate(scenes):
        end = scenes[i + 1]["idx"] if i + 1 < len(scenes) else len(paragraphs)
        body = []
        for text in paragraphs[scene["idx"] + 1 : end]:
            if text.startswith("人物"):
                continue
            if re.match(EPISODE_RE, text):
                continue
            body.append(text)
        scene["body"] = body
        scene["body_text"] = "\n".join(body)
    augment_visible_people(scenes)
    return scenes


def pdf_lines(path: Path) -> list[str]:
    if pdfplumber is None:
        raise RuntimeError("pdfplumber is required to parse PDF scripts")

    lines: list[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                extra_attrs=["fontname", "size", "non_stroking_color"],
                keep_blank_chars=False,
            )
            # Many script PDFs carry diagonal watermarks. These are usually
            # oversized gray words or a different watermark font; skip them
            # without deleting real Chinese characters from the script text.
            # Preserve real script text by default. Font names and large type are
            # not reliable watermark signals across Chinese screenplay exports.
            # Repeated watermark/header cleanup must be source-specific and
            # visually checked instead of globally discarding words here.
            filtered = words
            groups: list[dict] = []
            for word in filtered:
                for group in groups:
                    if abs(group["top"] - word["top"]) < 3:
                        group["words"].append(word)
                        group["top"] = (group["top"] + word["top"]) / 2
                        break
                else:
                    groups.append({"top": word["top"], "words": [word]})
            for group in sorted(groups, key=lambda item: item["top"]):
                text = "".join(
                    word["text"] for word in sorted(group["words"], key=lambda item: item["x0"])
                ).strip()
                if text:
                    lines.append(text)
    return lines


def is_pdf_heading(text: str) -> bool:
    normalized = text.replace(" ", "")
    full = r"^\d+\.\s*.+?(?:夜/日|日/夜|日|夜|清晨|上午|下午|傍晚|深夜|凌晨)(?:内外|内|外)$"
    malformed = r"^\d+\.\s*(?:（?闪回）?|（?闪现）?)?(?:一组镜头|手机屏幕|蒙太奇|转场).*$"
    return bool(re.match(full, normalized) or re.match(malformed, normalized))


def parse_pdf_heading(
    text: str,
    episode: int | str | None,
    scene_count: dict[int | str, int],
    last_day: str,
    last_inside_outside: str,
) -> tuple[str, str, str, str, str, str] | None:
    normalized = text.replace(" ", "")
    match = re.match(
        r"^(\d+)\.\s*(.+?)(夜/日|日/夜|日|夜|清晨|上午|下午|傍晚|深夜|凌晨)(内外|内|外)$",
        normalized,
    )
    warning = ""
    if match:
        source_num, location, day, inside_outside = match.groups()
        day = normalize_day(day)
    else:
        match = re.match(r"^(\d+)\.\s*(.+)$", normalized)
        if not match or not any(keyword in match.group(2) for keyword in ["一组镜头", "手机屏幕", "蒙太奇", "转场"]):
            return None
        source_num, location = match.groups()
        day = last_day if last_day in {"日", "夜"} else "需确认"
        inside_outside = "内外" if any(keyword in location for keyword in ["一组镜头", "蒙太奇", "转场"]) else (
            last_inside_outside if last_inside_outside in {"内", "外", "内外"} else "需确认"
        )
        warning = "原文未标日夜内外，需确认"

    if episode is None:
        return None
    scene_count[episode] = scene_count.get(episode, 0) + 1
    scene_id = f"{episode}-{scene_count[episode]}" if episode != "番外" else f"番外-{scene_count[episode]}"
    if scene_id.split("-")[-1] != source_num and episode != "番外":
        warning = "；".join([part for part in [warning, f"原文场号{source_num}"] if part])
    return scene_id, day, inside_outside, location.strip(), source_num, warning


def collect_pdf_people_line(lines: list[str], start: int, end: int) -> str:
    people_line = ""
    for idx in range(start + 1, min(end, start + 8)):
        text = lines[idx]
        if text.startswith("人物"):
            people_line = text
            cursor = idx + 1
            while cursor < min(end, idx + 4):
                next_text = lines[cursor].strip()
                if is_pdf_heading(next_text) or re.match(EPISODE_RE, next_text):
                    break
                if next_text.startswith("番外篇"):
                    break
                if people_line.endswith(("、", "，", ",")) or (
                    len(next_text) <= 16
                    and "：" not in next_text
                    and ":" not in next_text
                    and not re.search(r"[。！？；;]", next_text)
                ):
                    people_line += next_text
                    cursor += 1
                    continue
                break
            break
        if is_pdf_heading(text) or re.match(EPISODE_RE, text):
            break
    return people_line


def parse_pdf(path: Path) -> list[dict]:
    lines = pdf_lines(path)
    episode: int | str | None = None
    scene_count: dict[int | str, int] = {}
    scenes: list[dict] = []
    last_day = ""
    last_inside_outside = ""

    for idx, text in enumerate(lines):
        ep_match = re.match(EPISODE_RE, text.replace(" ", ""))
        if ep_match:
            episode = cn_episode_to_int(ep_match.group(1))
            if episode is not None:
                scene_count.setdefault(episode, 0)
            continue
        if text.replace(" ", "").startswith("番外篇"):
            episode = "番外"
            scene_count.setdefault(episode, 0)
            continue
        if not is_pdf_heading(text):
            continue
        parsed = parse_pdf_heading(text, episode, scene_count, last_day, last_inside_outside)
        if not parsed:
            continue
        scene_id, day, inside_outside, location, source_num, warning = parsed
        last_day = day
        last_inside_outside = inside_outside
        scenes.append(
            {
                "source": path.name,
                "idx": idx,
                "heading": text,
                "scene": scene_id,
                "source_num": source_num,
                "day": day,
                "inside_outside": inside_outside,
                "location": location,
                "people_raw": "",
                "people": [],
                "warning": warning,
                "heading_note": "",
            }
        )

    for scene_index, scene in enumerate(scenes):
        end = scenes[scene_index + 1]["idx"] if scene_index + 1 < len(scenes) else len(lines)
        people_line = collect_pdf_people_line(lines, scene["idx"], end)
        scene["people_raw"] = people_line
        scene["people"] = split_people(people_line)
        body = []
        for text in lines[scene["idx"] + 1 : end]:
            if text.startswith("人物"):
                continue
            if re.match(EPISODE_RE, text.replace(" ", "")):
                continue
            if text.replace(" ", "").startswith("番外篇"):
                continue
            body.append(text)
        scene["body"] = body
        scene["body_text"] = "\n".join(body)
    augment_visible_people(scenes)
    return scenes


def parse_script(path: Path) -> list[dict]:
    if path.suffix.lower() == ".pdf":
        return parse_pdf(path)
    if path.suffix.lower() in {".docx", ".docm"}:
        return parse_docx(path)
    if path.suffix.lower() == ".doc":
        if shutil.which("textutil") is None:
            raise RuntimeError("textutil is required to convert legacy .doc scripts")
        with tempfile.TemporaryDirectory(prefix="shunchangbiao-doc-") as tmpdir:
            converted = Path(tmpdir) / f"{path.stem}.docx"
            subprocess.run(
                ["textutil", "-convert", "docx", "-output", str(converted), str(path)],
                check=True,
                capture_output=True,
                text=True,
            )
            scenes = parse_docx(converted)
        for scene in scenes:
            scene["source"] = path.name
        return scenes
    raise ValueError(f"Unsupported input format: {path.suffix}. Expected .doc, .docx, .docm, or .pdf.")


def ordered_unique(items: list[str]) -> list[str]:
    seen = OrderedDict()
    for item in items:
        item = item.strip()
        if item:
            seen[item] = True
    return list(seen.keys())


def is_group_name(name: str) -> bool:
    if is_voice_name(name):
        return False
    if re.search(r"(?:甲|乙|丙|丁|\d+)$", name):
        return False
    return any(hint in name for hint in GROUP_HINTS)


def is_voice_name(name: str) -> bool:
    return bool(re.search(r"(?:VO|OS|画外音|电话声|广播声)$", name.strip(), flags=re.IGNORECASE))


def clean_voice_name(name: str) -> str:
    return re.sub(r"(?:VO|OS|画外音|电话声|广播声)$", "", name.strip(), flags=re.IGNORECASE).strip()


def voice_note(name: str) -> str:
    stripped = name.strip()
    suffix_match = re.search(r"(VO|OS|画外音|电话声|广播声)$", stripped, flags=re.IGNORECASE)
    if not suffix_match:
        return stripped
    suffix = suffix_match.group(1)
    if suffix.lower() == "vo":
        label = "VO"
    elif suffix.lower() == "os":
        label = "OS"
    else:
        label = suffix
    base = clean_voice_name(stripped)
    return f"{base}{label}" if base else label


def infer_roles(scenes: list[dict]) -> list[tuple[str, str]]:
    names = []
    for scene in scenes:
        for name in scene["people"]:
            if not is_group_name(name) and not is_voice_name(name):
                names.append(name)
    return [(name, default_mark(name)) for name in ordered_unique(names)]


def default_mark(role: str) -> str:
    for prefix, mark in [("幼年", "幼"), ("年幼", "幼"), ("少年", "少"), ("青年", "青"), ("老年", "老")]:
        if role.startswith(prefix):
            return mark
    if role.endswith("老爷子"):
        return "爷"
    if role.endswith("助理"):
        return role[0]
    if role.startswith("小") and len(role) >= 2:
        return role[-1]
    cleaned = re.sub(r"[（）()青年]", "", role)
    return cleaned[-1] if cleaned else role[-1]


def short_summary(scene: dict) -> str:
    fragments = []
    for line in scene["body"]:
        line = re.sub(r"^△", "", line).strip()
        line = re.sub(r"【[^】]*】", "", line)
        line = re.sub(r"[（(].*?[）)]", "", line)
        line = re.sub(r"^\w+[:：；;]", "", line).strip()
        line = re.sub(r"[，。！？、\s]+", "", line)
        if line:
            fragments.append(line)
        candidate = "".join(fragments)
        if len(candidate) >= 12:
            return candidate[:18]
    if fragments:
        return "".join(fragments)[:18]
    return "本场拍摄内容需要人工补写"


def collect_groups(scene: dict) -> str:
    groups = []
    text = scene["people_raw"] + "\n" + scene["body_text"]
    for person in scene["people"]:
        if is_group_name(person):
            groups.append(re.sub(r"(若干|众|[Nn]|\d+)$", "", person))
    for group_name in re.findall(
        r"机械武装队员|武装肃清队员|治安局肃清队员|治安局队员|厂区稽查队员|百乐门舞女|夜班劳工|流民",
        text,
    ):
        groups.append(group_name)
    if re.search(r"众人|所有人|全场|会议室里的人|各位", text):
        groups.append("会议人员" if "会议" in scene["location"] else "众人")
    if re.search(r"应聘者|面试者", text):
        groups.append("应聘者")
    if "保镖" in text:
        groups.append("保镖")
    if re.search(r"员工|同事|职员", text):
        groups.append("公司员工")
    return "、".join(ordered_unique(groups))


def collect_props(scene: dict) -> str:
    text = scene["body_text"]
    props = []
    for keyword, label in PROP_KEYWORDS:
        if keyword == "刀" and not re.search(r"持刀|抽出.*刀|举刀|刀刃|刀身|刀尖|短刀|菜刀|水果刀|美工刀|手中.*刀|刀.*砍|刀.*刺", text):
            continue
        for line in scene.get("body", []):
            if keyword not in line:
                continue
            if label in {"椅子", "病床", "花瓶", "水桶", "冷水", "菜肴", "饮料", "咖啡", "糕点"} and not prop_is_used(line, keyword):
                continue
            props.append(label)
            break
    if "头发" in text and re.search(r"拔了几根|扯了几根|亲子鉴定", text):
        props.append("头发样本")
    props = ordered_unique(props)
    if "水果刀" in props and "刀具" in props:
        props.remove("刀具")
    if "美工刀" in props and "刀具" in props:
        props.remove("刀具")
    if "子弹/枪械效果" in props and "枪械效果" in props:
        props.remove("枪械效果")
    if "手机" in props and "手机/电话" in props:
        props.remove("手机/电话")
    if "沈家纹路金属残片" in props and "金属残片" in props:
        props.remove("金属残片")
    if "全息光屏画面" in props and "光屏画面" in props:
        props.remove("光屏画面")
    return "、".join(props)


def prop_is_used(line: str, keyword: str) -> bool:
    action = r"拿|握|递|交|掏|打开|查看|使用|举|砍|刺|砸|推|坐|躺|拾|藏|收起|摊开|拆|戴|喝|吃|靠|绑|切|划|插|贴|搬|踢|泼|拔|扯"
    escaped = re.escape(keyword)
    return bool(
        re.search(rf"(?:{action}).{{0,8}}{escaped}", line)
        or re.search(rf"{escaped}.{{0,4}}(?:被|翻|倒|碎|破|移动)", line)
    )


def collect_costume(scene: dict) -> str:
    people = [person for person in scene.get("people", []) if not is_voice_name(person)]
    entries = []
    assigned_hints = []
    for person in people:
        role_lines = [line for line in scene.get("body", []) if person in line]
        role_hints = costume_hints_for_text("\n".join(role_lines))
        if "幼年" in person:
            role_hints.insert(0, "幼年造型")
        elif "青年" in person:
            role_hints.insert(0, "青年造型")
        elif "老年" in person:
            role_hints.insert(0, "老年造型")
        role_hints = ordered_unique(role_hints)
        if role_hints:
            entries.append(f"{person}：{'、'.join(role_hints)}")
            assigned_hints.extend(role_hints)

    all_hints = costume_hints_for_text(scene.get("body_text", ""))
    unassigned = [hint for hint in all_hints if hint not in assigned_hints]
    if unassigned:
        named_people = [person for person in people if not is_group_name(person)]
        if len(named_people) == 1 and not any(entry.startswith(f"{named_people[0]}：") for entry in entries):
            entries.append(f"{named_people[0]}：{'、'.join(unassigned)}")
        else:
            entries.append(f"未明确角色：{'、'.join(unassigned)}（需确认）")

    heading_note = scene.get("heading_note", "")
    if heading_note and re.search(r"接\s*\d+[-—–]\d+|同\s*\d+[-—–]\d+", heading_note):
        entries.append(f"连戏：{heading_note}")
    return "；".join(entries)


def costume_hints_for_text(text: str) -> list[str]:
    rules = [
        (r"怀孕|孕吐|反胃|摸了摸肚子", "孕期状态"),
        (r"湿透|滴水|泼冷水|浑身冰冷|湿发|湿衣", "湿发湿衣"),
        (r"灼伤|烧伤", "灼伤妆"),
        (r"包扎|受伤|划伤|伤口|带血|乌青|红肿", "伤口/血妆"),
        (r"昏迷|苍白|惨白|虚弱|憔悴|发烧|嘴唇发紫", "病弱妆"),
        (r"冷汗|汗水|汗湿", "冷汗妆效"),
        (r"义肢.*(?:破损|开裂|漏电)|(?:破损|开裂|漏电).*义肢", "机械义肢破损状态"),
        (r"衣衫不整", "衣衫不整"),
        (r"西装|高定", "西装/正装"),
        (r"制服|军装|警服", "制服"),
        (r"职业套装", "职业套装"),
        (r"旗袍", "旗袍"),
        (r"礼服", "礼服"),
        (r"睡衣|睡裙", "睡衣"),
        (r"连衣裙", "连衣裙"),
        (r"洗得发白|旧外套|破旧衣服", "朴素旧衣"),
        (r"保洁|清洁工", "保洁服/朴素造型"),
        (r"衣衫褴褛|浑身脏污", "狼狈脏污造型"),
        (r"盘发", "盘发"),
        (r"散发|头发凌乱", "散发/凌乱发型"),
    ]
    return [label for pattern, label in rules if re.search(pattern, text)]


def collect_notes(scene: dict) -> str:
    text = scene["body_text"]
    notes = []
    if "字幕" in text:
        notes.append("含字幕")
    voice_mentions = re.findall(
        r"(?im)^(?:△\s*)?([一-鿿A-Za-z0-9]+)\s*(VO|OS)\s*(?:[（(][^）)]*[）)])?\s*[:：]",
        text,
    )
    if voice_mentions:
        notes.extend(f"{speaker}{kind.upper()}" for speaker, kind in voice_mentions)
    elif re.search(r"VO|vo|画外音", text):
        notes.append("含VO")
    if re.search(r"闪回|回忆", text):
        notes.append("闪回/回忆")
    if re.search(r"接吻|亲吻|强吻|乱亲|床上|压在|抱住.*亲|衣衫不整|睡裙|钻进被窝", text):
        notes.append("亲密戏注意尺度")
    if re.search(r"扭打|持刀|水果刀|美工刀|踹开|砸|扑倒|推倒|滚远|撞在地上|拽|拉扯|摔", text):
        notes.append("动作戏注意安全")
    if "接" in scene["location"]:
        notes.append(re.sub(r".*?(接[^）)]*).*", r"\1", scene["location"]))
    if "一同拍摄" in scene["location"]:
        notes.append("一同拍摄")
    if scene.get("warning"):
        notes.append(scene["warning"])
    if scene.get("heading_note"):
        notes.append(scene["heading_note"])
    for person in scene.get("people", []):
        if is_voice_name(person):
            notes.append(voice_note(person))
    return "；".join(ordered_unique(notes))


def vertical(text: str) -> str:
    return "\n".join(text)


def load_config(path: Path | None) -> dict:
    if not path:
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_roles(scenes: list[dict], config: dict) -> list[tuple[str, str]]:
    role_config = config.get("roles")
    if role_config:
        return [(item["name"], item.get("mark") or default_mark(item["name"])) for item in role_config]
    return infer_roles(scenes)


def scene_review(config: dict, scene_id: str) -> dict:
    review = dict(config.get("scenes", {}).get(scene_id, {}))
    legacy_summary = config.get("summaries", {}).get(scene_id)
    if legacy_summary is not None and "summary" not in review:
        review["summary"] = legacy_summary
    return review


def reviewed_or_default(review: dict, key: str, default: str) -> str:
    return review[key] if key in review else default


def validation_issues(
    scenes: list[dict],
    config: dict,
    roles: list[tuple[str, str]],
    strict: bool,
) -> list[str]:
    issues = []
    scene_ids = [scene["scene"] for scene in scenes]
    duplicates = ordered_unique(scene_id for scene_id in scene_ids if scene_ids.count(scene_id) > 1)
    if not scenes:
        issues.append("未识别到任何场次")
    if duplicates:
        issues.append(f"场号重复：{'、'.join(duplicates)}")

    marks = [mark for _, mark in roles if mark]
    duplicate_marks = ordered_unique(mark for mark in marks if marks.count(mark) > 1)
    if duplicate_marks:
        issues.append(f"角色简称重复：{'、'.join(duplicate_marks)}")
    invalid_roles = [name for name, _ in roles if is_group_name(name) or is_voice_name(name)]
    if invalid_roles:
        issues.append(f"群演或VO误入角色列：{'、'.join(invalid_roles)}")

    if not strict:
        return issues

    for scene in scenes:
        scene_id = scene["scene"]
        review = scene_review(config, scene_id)
        if review.get("reviewed") is not True:
            issues.append(f"{scene_id}：未完成逐场复核")
            continue
        required_fields = {"summary", "groups", "costume", "props", "notes"}
        missing_fields = sorted(required_fields - set(review))
        if missing_fields:
            issues.append(f"{scene_id}：复核字段不完整（缺少{'、'.join(missing_fields)}）")
        summary = str(review.get("summary", "")).strip()
        if not summary:
            issues.append(f"{scene_id}：拍摄内容未填写")
        else:
            compact_length = len(re.sub(r"[\s，。！？；、,.;!?]", "", summary))
            special_scene = bool(re.search(r"空镜|一组镜头|蒙太奇|手机屏幕|转场", scene.get("location", "") + scene.get("heading", "")))
            minimum = 4 if special_scene else 12
            if not minimum <= compact_length <= 18:
                issues.append(f"{scene_id}：拍摄内容应为{minimum}-18个汉字")
        if scene.get("day") == "需确认" or scene.get("inside_outside") == "需确认":
            notes = str(reviewed_or_default(review, "notes", collect_notes(scene)))
            if "需确认" not in notes:
                issues.append(f"{scene_id}：日夜/内外缺失但备注未标记需确认")
    return issues


def build_workbook(source: Path, scenes: list[dict], output: Path, config: dict) -> None:
    roles = resolve_roles(scenes, config)

    wb = Workbook()
    ws = wb.active
    ws.title = "顺场表"
    headers = ["拍摄顺序", "实际拍摄场地", "场次", "剧本中场景", "拍摄内容"]
    headers += ["日\n/\n夜", "内\n/\n外", "页\n数"]
    headers += [vertical(role) for role, _ in roles]
    headers += ["群演", "梳化服提示", "道具提示", "备注/特殊道具/时间"]
    ws.append(headers)

    role_names = {role for role, _ in roles}
    for order, scene in enumerate(scenes, start=1):
        people = set(scene["people"])
        review = scene_review(config, scene["scene"])
        row = [
            order,
            "",
            scene["scene"],
            scene["location"],
            reviewed_or_default(review, "summary", short_summary(scene)),
            scene["day"],
            scene["inside_outside"],
            "",
        ]
        for role, mark in roles:
            row.append(mark if role in people and role in role_names else "")
        row += [
            reviewed_or_default(review, "groups", collect_groups(scene)),
            reviewed_or_default(review, "costume", collect_costume(scene)),
            reviewed_or_default(review, "props", collect_props(scene)),
            reviewed_or_default(review, "notes", collect_notes(scene)),
        ]
        ws.append(row)

    style_sheet(ws, len(headers), len(scenes), len(roles))
    wb.save(output)


def style_sheet(ws, max_col: int, scene_count: int, role_count: int) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E5F")
    header_font = Font(name="Microsoft YaHei", color="FFFFFF", bold=True, size=10)
    body_font = Font(name="Microsoft YaHei", color="222222", size=10)
    alt_fill = PatternFill("solid", fgColor="F4F8FA")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    thin = Side(style="thin", color="B8C7CE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "I2"
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{scene_count + 1}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"
    ws.print_area = f"A1:{get_column_letter(max_col)}{scene_count + 1}"

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = max(82, min(140, 12 * max(6, role_count)))

    for row in range(2, scene_count + 2):
        fill = alt_fill if row % 2 == 0 else white_fill
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = fill
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col in [4, 5, max_col - 3, max_col - 2, max_col - 1, max_col]:
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        text_columns = [
            (4, 23),
            (5, 20),
            (max_col - 3, 15),
            (max_col - 2, 20),
            (max_col - 1, 22),
            (max_col, 26),
        ]
        estimated_lines = 1
        for col, width in text_columns:
            value = str(ws.cell(row=row, column=col).value or "")
            line_count = sum(max(1, math.ceil(len(part) / width)) for part in value.split("\n"))
            estimated_lines = max(estimated_lines, line_count)
        ws.row_dimensions[row].height = min(108, max(36, 20 * estimated_lines + 8))

    for col, width in {1: 8, 2: 14, 3: 9, 4: 23, 5: 20}.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    for col in range(6, 9):
        ws.column_dimensions[get_column_letter(col)].width = 4.2
    role_start = 9
    role_end = role_start + role_count - 1
    for col in range(role_start, role_end + 1):
        ws.column_dimensions[get_column_letter(col)].width = 4.2
    ws.column_dimensions[get_column_letter(role_end + 1)].width = 15
    ws.column_dimensions[get_column_letter(role_end + 2)].width = 20
    ws.column_dimensions[get_column_letter(role_end + 3)].width = 22
    ws.column_dimensions[get_column_letter(role_end + 4)].width = 26


def title_from_path(path: Path) -> str:
    name = path.stem
    match = re.search(r"《([^》]+)》", name)
    return match.group(1) if match else name


def dump_json(path: Path, all_scenes: dict[str, list[dict]]) -> None:
    serializable = {}
    for key, scenes in all_scenes.items():
        serializable[key] = [
            {k: v for k, v in scene.items() if k not in {"body_text"}}
            for scene in scenes
        ]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(serializable, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build short-drama 顺场表 Excel files from DOCX/PDF scripts.")
    parser.add_argument("scripts", nargs="+", type=Path)
    parser.add_argument("--out-dir", type=Path, default=Path.cwd())
    parser.add_argument("--config", type=Path)
    parser.add_argument("--dump-json", type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--strict", action="store_true", help="Block delivery until every scene is reviewed and validation passes.")
    args = parser.parse_args()

    config = load_config(args.config)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    all_scenes = {}
    reports = {}

    for source in args.scripts:
        scenes = parse_script(source)
        all_scenes[source.name] = scenes
        output = args.out_dir / f"顺场表-{title_from_path(source)}.xlsx"
        per_file_config = config.get(source.name, config)
        roles = resolve_roles(scenes, per_file_config)
        issues = validation_issues(scenes, per_file_config, roles, strict=args.strict)
        reports[source.name] = {
            "scene_count": len(scenes),
            "issues": issues,
            "strict": args.strict,
            "draft_valid": not issues,
            "delivery_allowed": args.strict and not issues,
        }
        if issues:
            print(f"{source.name}: validation issues")
            for issue in issues:
                print(f"  - {issue}")
            if args.strict:
                if args.report:
                    args.report.parent.mkdir(parents=True, exist_ok=True)
                    args.report.write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
                raise SystemExit(2)
        build_workbook(source, scenes, output, per_file_config)
        warnings = [f"{s['scene']}:{s['warning']}" for s in scenes if s.get("warning")]
        print(f"{source.name}: {len(scenes)} scenes -> {output}")
        if warnings:
            print("  warnings:", ", ".join(warnings))

    if args.dump_json:
        dump_json(args.dump_json, all_scenes)
        print(f"scene dump -> {args.dump_json}")
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"validation report -> {args.report}")


if __name__ == "__main__":
    main()
