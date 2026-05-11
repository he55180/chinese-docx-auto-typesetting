# -*- coding: utf-8 -*-
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from format_expert import (
    to_chinese_numeral, get_paragraph_type, strip_numbering,
    is_main_title, is_likely_heading, should_add_prefix,
    split_at_first_punct, is_heading_with_body, NumberingTracker
)


class TestChineseNumeral:
    def test_basic(self):
        assert to_chinese_numeral(0) == ''
        assert to_chinese_numeral(1) == '一'
        assert to_chinese_numeral(5) == '五'
        assert to_chinese_numeral(9) == '九'

    def test_tens(self):
        assert to_chinese_numeral(10) == '十'
        assert to_chinese_numeral(11) == '十一'
        assert to_chinese_numeral(15) == '十五'
        assert to_chinese_numeral(19) == '十九'

    def test_twenties_and_beyond(self):
        assert to_chinese_numeral(20) == '二十'
        assert to_chinese_numeral(25) == '二十五'
        assert to_chinese_numeral(30) == '三十'
        assert to_chinese_numeral(99) == '九十九'

    def test_out_of_range(self):
        assert to_chinese_numeral(100) == ''
        assert to_chinese_numeral(-1) == ''


class TestParagraphType:
    def test_empty(self):
        assert get_paragraph_type('') == 'empty'

    def test_chapter(self):
        assert get_paragraph_type('第一章：项目概述') == 'chapter'
        assert get_paragraph_type('第十一章 安全管理') == 'chapter'

    def test_section(self):
        assert get_paragraph_type('第一节 检查范围') == 'section'

    def test_level1(self):
        assert get_paragraph_type('一、项目概述') == 'level1'
        assert get_paragraph_type('十、总结') == 'level1'

    def test_level1_shi(self):
        assert get_paragraph_type('一是加强安全管理') == 'level1_shi'
        assert get_paragraph_type('二是提升质量水平') == 'level1_shi'

    def test_level2(self):
        assert get_paragraph_type('（一）检查范围') == 'level2'

    def test_level3(self):
        assert get_paragraph_type('1. 项目背景') == 'level3'

    def test_body(self):
        assert get_paragraph_type('这是一段普通的正文内容，用于测试正文识别功能。') == 'body'


class TestStripNumbering:
    def test_strip_chapter(self):
        assert strip_numbering('第一章：项目概述') == '项目概述'
        assert strip_numbering('第十一章 安全管理') == '安全管理'

    def test_strip_level1(self):
        assert strip_numbering('一、项目概述') == '项目概述'

    def test_strip_level2(self):
        assert strip_numbering('（一）检查范围') == '检查范围'

    def test_no_numbering(self):
        assert strip_numbering('正文段落没有编号') == '正文段落没有编号'


class TestMainTitle:
    def test_keyword_match(self):
        assert is_main_title('关于安全生产工作的通知', True) is True

    def test_not_first_para(self):
        assert is_main_title('关于通知', False) is False

    def test_empty(self):
        assert is_main_title('', True) is False

    def test_red_head_rejection(self):
        assert is_main_title('某公司〔2026〕12号文件', True) is False
        assert is_main_title('某公司2026年12号', True) is False

    def test_year_in_title_not_rejected(self):
        assert is_main_title('中国港湾十五五发展规划报告（2026-2030年）', True) is True

    def test_too_long(self):
        long_text = '这是一段超过五十个字的非常长的文本' * 3
        assert is_main_title(long_text, True) is False


class TestShouldAddPrefix:
    def test_existing_chapter(self):
        assert should_add_prefix('第一章', 'chapter') is False

    def test_missing_prefix(self):
        assert should_add_prefix('项目概述', 'level1') is True

    def test_existing_section(self):
        assert should_add_prefix('第一节 概述', 'section') is False


class TestSplitAtFirstPunct:
    def test_period_split(self):
        h, b = split_at_first_punct('项目概述。项目位于东非区域。')
        assert h == '项目概述。'
        assert b == '项目位于东非区域。'

    def test_colon_split_compact(self):
        h, b = split_at_first_punct('市场规模快速增长：新签合同额年均增长15%。')
        assert h == '市场规模快速增长：'
        assert b == '新签合同额年均增长15%。'

    def test_heading_preserves_subtitle(self):
        h, b = split_at_first_punct('项目名：副标题名称。正文从这里开始', is_heading=True)
        assert h == '项目名：副标题名称。'
        assert b == '正文从这里开始'

    def test_no_punct(self):
        h, b = split_at_first_punct('项目概述')
        assert h == '项目概述'
        assert b == ''

    def test_no_comma_split(self):
        h, b = split_at_first_punct('数字化、智能化建造')
        assert h == '数字化、智能化建造'
        assert b == ''


class TestIsHeadingWithBody:
    def test_level1(self):
        assert is_heading_with_body('一、项目概述') is True

    def test_level2(self):
        assert is_heading_with_body('（一）检查范围') is True

    def test_numbered(self):
        assert is_heading_with_body('1. 项目背景') is True

    def test_body_not_heading(self):
        assert is_heading_with_body('这是一段正文') is False

    def test_body_with_colon_not_heading(self):
        assert is_heading_with_body('市场规模快速增长：新签合同额年均增长。') is False


class TestNumberingTracker:
    def test_basic(self):
        t = NumberingTracker()
        assert t.next('chapter') == 1
        assert t.next('section') == 1
        assert t.next('level1') == 1

    def test_reset_on_higher(self):
        t = NumberingTracker()
        t.next('chapter')  # ch1
        t.next('section')  # sec1
        t.next('level1')   # l1-1
        t.next('level1')   # l1-2
        assert t.next('section') == 2  # new section resets level1
        assert t.next('level1') == 1   # should be 1, not 3

    def test_chapter_resets_all(self):
        t = NumberingTracker()
        t.next('chapter')
        t.next('section')
        t.next('section')
        t.next('level1')
        t.next('level1')
        assert t.next('chapter') == 2
        assert t.counters['section'] == 0
        assert t.counters['level1'] == 0
