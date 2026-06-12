"""Tuomin 测试套件"""
import os
import tempfile

import pytest
from openpyxl import Workbook

from tuomin import (
    MaskingEngine,
    col_letter_to_index,
    parse_col_spec,
    read_sheet,
    replace_address,
    replace_bankcard,
    replace_email,
    replace_idcard,
    replace_name,
    replace_phone,
    resolve_output_path,
    reconstruct_name,
    reconstruct_phone,
    reconstruct_idcard,
    reconstruct_email,
    reconstruct_bankcard,
    reconstruct_address,
    write_xlsx,
)


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------


class TestColLetterToIndex:
    def test_single(self):
        assert col_letter_to_index("A") == 0
        assert col_letter_to_index("B") == 1
        assert col_letter_to_index("Z") == 25

    def test_double(self):
        assert col_letter_to_index("AA") == 26
        assert col_letter_to_index("AB") == 27

    def test_case_insensitive(self):
        assert col_letter_to_index("a") == 0
        assert col_letter_to_index("z") == 25

    def test_invalid(self):
        with pytest.raises(ValueError):
            col_letter_to_index("1")


class TestParseColSpec:
    def test_valid(self):
        idx, ft, m = parse_col_spec("C:phone:replace")
        assert idx == 2
        assert ft == "phone"
        assert m == "replace"

    def test_whitespace(self):
        idx, ft, m = parse_col_spec(" B : NAME : Reconstruct ")
        assert idx == 1
        assert ft == "name"
        assert m == "reconstruct"

    def test_bad_format(self):
        with pytest.raises(ValueError):
            parse_col_spec("A:phone")

    def test_bad_type(self):
        with pytest.raises(ValueError):
            parse_col_spec("A:foo:replace")

    def test_bad_method(self):
        with pytest.raises(ValueError):
            parse_col_spec("A:phone:delete")


# ---------------------------------------------------------------------------
# 替换规则
# ---------------------------------------------------------------------------


class TestReplaceName:
    def test_two_char(self):
        assert replace_name("张三") == "张*"

    def test_three_char(self):
        assert replace_name("张三丰") == "张**"

    def test_empty(self):
        assert replace_name("") == ""


class TestReplacePhone:
    def test_normal(self):
        assert replace_phone("13812341234") == "138****1234"

    def test_with_spaces(self):
        assert replace_phone("138 1234 1234") == "138****1234"

    def test_short(self):
        assert replace_phone("12345") == "12345"

    def test_non_numeric(self):
        assert replace_phone("未知") == "未知"


class TestReplaceIdcard:
    def test_eighteen(self):
        result = replace_idcard("310101199001011234")
        assert result == "310***********1234"

    def test_short(self):
        assert replace_idcard("1234567") == "1234567"


class TestReplaceEmail:
    def test_normal(self):
        assert replace_email("zhangsan@qq.com") == "z*******@qq.com"

    def test_single_char_local(self):
        assert replace_email("a@b.com") == "a@b.com"

    def test_no_at(self):
        assert replace_email("notanemail") == "notanemail"


class TestReplaceBankcard:
    def test_sixteen(self):
        assert replace_bankcard("6222021234565678") == "6222********5678"

    def test_short(self):
        assert replace_bankcard("12345678") == "12345678"


class TestReplaceAddress:
    def test_normal(self):
        result = replace_address("上海市浦东新区张江路88号")
        assert result.startswith("上海市浦东新区")
        assert "*" in result

    def test_province_city(self):
        result = replace_address("广东省深圳市南山区科技园路12号")
        assert result.startswith("广东省深圳市")

    def test_no_province(self):
        assert replace_address("某某地方") == "某某地方"


# ---------------------------------------------------------------------------
# 重构规则
# ---------------------------------------------------------------------------


class TestReconstructName:
    def test_format(self):
        result = reconstruct_name("anything")
        assert len(result) >= 3
        assert result[0] in "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮卞齐康伍余元卜顾孟平黄和穆萧尹"

    def test_ignores_original(self):
        r1 = reconstruct_name("张三")
        r2 = reconstruct_name("李四")
        # 不一定不同（随机），但格式应该一致
        assert len(r1) >= 3
        assert len(r2) >= 3


class TestReconstructPhone:
    def test_starts_110(self):
        result = reconstruct_phone("anything")
        assert result.startswith("110")
        assert len(result) == 11
        assert result.isdigit()


class TestReconstructIdcard:
    def test_starts_999(self):
        result = reconstruct_idcard("anything")
        assert result.startswith("999")
        assert len(result) == 18

    def test_check_digit_valid(self):
        from tuomin import _idcard_check_digit
        result = reconstruct_idcard("anything")
        expected = _idcard_check_digit(result[:17])
        assert result[-1] == expected


class TestReconstructEmail:
    def test_format(self):
        result = reconstruct_email("anything")
        assert result.endswith("@notexist.invalid")
        local = result.split("@")[0]
        assert len(local) == 5


class TestReconstructBankcard:
    def test_starts_0000(self):
        result = reconstruct_bankcard("anything")
        assert result.startswith("0000")
        assert len(result) == 16


class TestReconstructAddress:
    def test_format(self):
        result = reconstruct_address("anything")
        assert "虚构路" in result
        assert result.endswith("号")


# ---------------------------------------------------------------------------
# MaskingEngine 一致性
# ---------------------------------------------------------------------------


class TestMaskingEngine:
    def test_consistency(self):
        engine = MaskingEngine()
        r1 = engine.mask("13812341234", "phone", "reconstruct")
        r2 = engine.mask("13812341234", "phone", "reconstruct")
        assert r1 == r2

    def test_different_values_different_results(self):
        engine = MaskingEngine()
        r1 = engine.mask("13812341234", "phone", "reconstruct")
        r2 = engine.mask("15987654321", "phone", "reconstruct")
        assert r1 != r2

    def test_none_passthrough(self):
        engine = MaskingEngine()
        assert engine.mask(None, "phone", "replace") is None

    def test_empty_passthrough(self):
        engine = MaskingEngine()
        assert engine.mask("", "phone", "replace") == ""

    def test_overlay_reconstruct(self):
        engine = MaskingEngine()
        r1 = engine.mask("张**", "name", "reconstruct")
        assert "*" not in r1
        assert len(r1) >= 3

    def test_overlay_phone_reconstruct(self):
        engine = MaskingEngine()
        r1 = engine.mask("138****1234", "phone", "reconstruct")
        assert r1.startswith("110")
        assert len(r1) == 11


# ---------------------------------------------------------------------------
# 文件读写集成测试
# ---------------------------------------------------------------------------


class TestFileIO:
    def _create_test_xlsx(self, tmpdir):
        wb = Workbook()
        ws = wb.active
        ws.append(["序号", "姓名", "手机号"])
        ws.append([1, "张三", "13812341234"])
        ws.append([2, "李四", "15987654321"])
        path = os.path.join(tmpdir, "test.xlsx")
        wb.save(path)
        return path

    def test_read_xlsx(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = self._create_test_xlsx(tmpdir)
            rows, col_count, _ = read_sheet(path)
            assert col_count == 3
            assert len(rows) == 3
            assert rows[1][1] == "张三"

    def test_write_xlsx(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            rows = [["序号", "姓名"], ["1", "赵六"]]
            out = os.path.join(tmpdir, "out.xlsx")
            write_xlsx(rows, out)
            assert os.path.isfile(out)

    def test_end_to_end(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = self._create_test_xlsx(tmpdir)
            rows, col_count, _ = read_sheet(path)

            engine = MaskingEngine()
            for row in rows:
                row[1] = engine.mask(row[1], "name", "replace")
                row[2] = engine.mask(row[2], "phone", "replace")

            out = os.path.join(tmpdir, "result.xlsx")
            write_xlsx(rows, out)

            result_rows, _, _ = read_sheet(out)
            assert result_rows[1][1] == "张*"
            assert result_rows[1][2] == "138****1234"

    def test_resolve_output_default(self):
        result = resolve_output_path("/tmp/data.xlsx")
        assert result.endswith("_masked.xlsx")
        assert "data_masked.xlsx" in result

    def test_resolve_output_custom(self):
        result = resolve_output_path("/tmp/data.xlsx", "/tmp/out.xlsx")
        assert result.endswith("out.xlsx")
