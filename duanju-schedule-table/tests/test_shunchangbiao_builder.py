from __future__ import annotations

import sys
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

from docx import Document
from openpyxl import load_workbook


SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

import shunchangbiao_builder as builder  # noqa: E402


class SceneHeadingTests(unittest.TestCase):
    def test_numbered_empty_shot_is_always_a_scene_boundary(self) -> None:
        parsed = builder.parse_heading("1-1空镜", 1, {}, "", "")

        self.assertIsNotNone(parsed)
        self.assertEqual(parsed[:4], ("1-1", "需确认", "需确认", "空镜"))
        self.assertIn("日夜需确认", parsed[4])
        self.assertIn("内外需确认", parsed[4])

    def test_trailing_flashback_annotation_does_not_hide_scene(self) -> None:
        parsed = builder.parse_heading(
            "2-2 沈家旧宅小院  夜 外（闪回 同3-1）",
            2,
            {},
            "夜",
            "外",
        )

        self.assertIsNotNone(parsed)
        self.assertEqual(parsed[:4], ("2-2", "夜", "外", "沈家旧宅小院"))
        self.assertEqual(parsed[5], "闪回 同3-1")

    def test_chinese_episode_numbers_over_thirty_are_supported(self) -> None:
        self.assertEqual(builder.cn_episode_to_int("三十一"), 31)
        self.assertEqual(builder.cn_episode_to_int("八十"), 80)
        self.assertEqual(builder.cn_episode_to_int("一百零二"), 102)

    def test_docx_scene_bodies_stop_at_incomplete_headings(self) -> None:
        document = Document()
        for paragraph in [
            "正文",
            "第一集",
            "1-1空镜",
            "△城市外景。",
            "1-2 百乐门外空镜+下城区空镜",
            "△百乐门外。",
            "1-3 百乐门舞厅 夜 内",
            "人物：沈晚晴",
            "△沈晚晴走进舞厅。",
            "第二集",
            "2-1 贫民窟暗巷 夜 外",
            "人物：沈晚晴",
            "△沈晚晴拐入暗巷。",
            "2-2 沈家旧宅 夜 内（闪回）",
            "人物：沈父、幼年沈晚晴",
            "△沈父护住幼年沈晚晴。",
        ]:
            document.add_paragraph(paragraph)

        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "fixture.docx"
            document.save(path)
            scenes = builder.parse_docx(path)

        self.assertEqual([scene["scene"] for scene in scenes], ["1-1", "1-2", "1-3", "2-1", "2-2"])
        self.assertNotIn("百乐门外", scenes[0]["body_text"])
        self.assertNotIn("沈父", scenes[3]["body_text"])

    def test_docx_table_content_is_read_in_document_order(self) -> None:
        document = Document()
        document.add_paragraph("正文")
        document.add_paragraph("第一集")
        table = document.add_table(rows=4, cols=1)
        table.cell(0, 0).text = "1-1 办公室 日 内"
        table.cell(1, 0).text = "人物：陆锋"
        table.cell(2, 0).text = "△陆锋查看文件。"
        table.cell(3, 0).text = "1-2 走廊 日 内"

        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "table-script.docx"
            document.save(path)
            scenes = builder.parse_docx(path)

        self.assertEqual([scene["scene"] for scene in scenes], ["1-1", "1-2"])
        self.assertIn("陆锋查看文件", scenes[0]["body_text"])

    def test_legacy_doc_is_converted_and_parsed(self) -> None:
        if shutil.which("textutil") is None:
            self.skipTest("textutil is unavailable")
        document = Document()
        for paragraph in ["正文", "第一集", "1-1 办公室 日 内", "人物：陆锋", "△陆锋查看文件。"]:
            document.add_paragraph(paragraph)

        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = Path(tmpdir) / "legacy-source.docx"
            doc_path = Path(tmpdir) / "legacy-source.doc"
            document.save(docx_path)
            subprocess.run(
                ["textutil", "-convert", "doc", "-output", str(doc_path), str(docx_path)],
                check=True,
                capture_output=True,
                text=True,
            )
            scenes = builder.parse_script(doc_path)

        self.assertEqual([scene["scene"] for scene in scenes], ["1-1"])


class PeopleAndDepartmentTests(unittest.TestCase):
    def test_group_roles_and_voice_roles_do_not_become_actor_columns(self) -> None:
        self.assertTrue(builder.is_group_name("百乐门舞女"))
        self.assertTrue(builder.is_group_name("治安局队员众"))
        self.assertTrue(builder.is_group_name("厂区稽查队员"))
        self.assertTrue(builder.is_voice_name("巡逻员vo"))

        scenes = [
            {
                "people": ["沈晚晴", "百乐门舞女", "治安局队员众", "巡逻员vo"],
            }
        ]
        self.assertEqual(builder.infer_roles(scenes), [("沈晚晴", "晴")])

    def test_age_variant_prevents_duplicate_base_role_in_flashback(self) -> None:
        scenes = [
            {
                "people": ["沈父", "幼年沈晚晴"],
                "body": ["沈晚晴（虚弱）：爸爸！"],
                "heading_note": "闪回",
            },
            {
                "people": ["沈晚晴"],
                "body": ["△沈晚晴走进舞厅。"],
                "heading_note": "",
            },
        ]

        builder.augment_visible_people(scenes)

        self.assertEqual(scenes[0]["people"], ["沈父", "幼年沈晚晴"])
        self.assertEqual(builder.default_mark("沈晚晴"), "晴")
        self.assertEqual(builder.default_mark("幼年沈晚晴"), "幼")

    def test_costume_hints_are_attributed_to_roles(self) -> None:
        scene = {
            "people_raw": "人物：沈晚晴、陆锋",
            "people": ["沈晚晴", "陆锋"],
            "body": [
                "△沈晚晴右腿接驳处灼伤，额头冷汗。",
                "△陆锋身穿治安局制服。",
            ],
            "body_text": "△沈晚晴右腿接驳处灼伤，额头冷汗。\n△陆锋身穿治安局制服。",
        }

        hints = builder.collect_costume(scene)

        self.assertIn("沈晚晴：", hints)
        self.assertIn("陆锋：制服", hints)
        self.assertIn("灼伤", hints)

    def test_props_exclude_background_furniture_and_keep_story_props(self) -> None:
        scene = {
            "body": ["△房间里摆着椅子，沈晚晴收起口红，拿起酒杯，又将沈家纹路金属残片贴在全息光屏上。"],
            "body_text": "△房间里摆着椅子，沈晚晴收起口红，拿起酒杯，又将沈家纹路金属残片贴在全息光屏上。",
        }

        props = builder.collect_props(scene)

        self.assertIn("沈家纹路金属残片", props)
        self.assertIn("全息光屏画面", props)
        self.assertIn("口红", props)
        self.assertIn("酒杯", props)
        self.assertNotIn("椅子", props)

    def test_groups_are_detected_from_action_text_without_people_line(self) -> None:
        scene = {
            "location": "沈家旧宅小院",
            "people_raw": "",
            "people": [],
            "body_text": "△大批机械武装队员从院外破门突入，瞬间封锁小院。",
        }

        self.assertIn("机械武装队员", builder.collect_groups(scene))

    def test_voice_notes_preserve_vo_and_os_types(self) -> None:
        scene = {
            "location": "办公室",
            "people": ["巡逻员vo"],
            "body_text": "陆锋os：这件事没有那么简单。",
            "warning": "",
            "heading_note": "",
        }

        notes = builder.collect_notes(scene)

        self.assertIn("巡逻员VO", notes)
        self.assertIn("陆锋OS", notes)
        self.assertNotIn("陆锋VO", notes)


class DeliveryValidationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.scene = {
            "scene": "1-1",
            "location": "城市空镜",
            "day": "需确认",
            "inside_outside": "需确认",
            "heading_note": "",
            "warning": "日夜需确认；内外需确认",
            "people_raw": "",
            "people": [],
            "body": ["△城市空镜交代下城区环境。"],
            "body_text": "△城市空镜交代下城区环境。",
        }

    def test_strict_delivery_requires_every_scene_to_be_reviewed(self) -> None:
        issues = builder.validation_issues([self.scene], {}, [], strict=True)

        self.assertTrue(any("未完成逐场复核" in issue for issue in issues))

    def test_reviewed_scene_with_explicit_unknowns_can_pass_strict_delivery(self) -> None:
        config = {
            "scenes": {
                "1-1": {
                    "reviewed": True,
                    "summary": "城市空镜交代下城区夜间环境",
                    "groups": "",
                    "costume": "",
                    "props": "",
                    "notes": "日夜、内外需确认",
                }
            }
        }

        self.assertEqual(builder.validation_issues([self.scene], config, [], strict=True), [])

    def test_strict_delivery_requires_all_reviewed_fields_even_when_blank(self) -> None:
        config = {
            "scenes": {
                "1-1": {
                    "reviewed": True,
                    "summary": "城市空镜交代下城区夜间环境",
                    "notes": "日夜、内外需确认",
                }
            }
        }

        issues = builder.validation_issues([self.scene], config, [], strict=True)

        self.assertTrue(any("复核字段不完整" in issue for issue in issues))

    def test_duplicate_role_marks_are_rejected(self) -> None:
        roles = [("沈晚晴", "晴"), ("顾晚晴", "晴")]

        issues = builder.validation_issues([self.scene], {}, roles, strict=False)

        self.assertTrue(any("角色简称重复" in issue for issue in issues))

    def test_reviewed_fields_feed_the_unchanged_workbook_headers(self) -> None:
        config = {
            "roles": [{"name": "沈晚晴", "mark": "晴"}],
            "scenes": {
                "1-1": {
                    "reviewed": True,
                    "summary": "城市空镜交代下城区夜间环境",
                    "groups": "流民",
                    "costume": "流民：破旧衣服、脏污妆",
                    "props": "雾气特效设备",
                    "notes": "空镜；日夜、内外需确认；需要分别拍摄城市全景、百乐门外立面和下城区街道空镜",
                }
            },
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "schedule.xlsx"
            builder.build_workbook(Path("fixture.docx"), [self.scene], output, config)
            workbook = load_workbook(output)
            sheet = workbook["顺场表"]
            headers = [cell.value for cell in sheet[1]]
            values = [cell.value for cell in sheet[2]]
            freeze_panes = sheet.freeze_panes
            print_title_rows = sheet.print_title_rows
            row_height = sheet.row_dimensions[2].height
            workbook.close()

        self.assertEqual(
            headers,
            [
                "拍摄顺序",
                "实际拍摄场地",
                "场次",
                "剧本中场景",
                "拍摄内容",
                "日\n/\n夜",
                "内\n/\n外",
                "页\n数",
                "沈\n晚\n晴",
                "群演",
                "梳化服提示",
                "道具提示",
                "备注/特殊道具/时间",
            ],
        )
        self.assertEqual(values[4], "城市空镜交代下城区夜间环境")
        self.assertEqual(values[-4:-1], ["流民", "流民：破旧衣服、脏污妆", "雾气特效设备"])
        self.assertIn("需要分别拍摄城市全景", values[-1])
        self.assertEqual(freeze_panes, "I2")
        self.assertEqual(print_title_rows, "$1:$1")
        self.assertGreater(row_height, 36)


if __name__ == "__main__":
    unittest.main()
