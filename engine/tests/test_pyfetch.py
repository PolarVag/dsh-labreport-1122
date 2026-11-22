"""wheel 挑选逻辑的测试。

这里以前踩过两个真坑：① cp314-cp314（abi 既不等于 none 也不等于 abi3）被判成不兼容；
② 32 位的 win32 wheel 抢过 win_amd64，装完 `from lxml import etree` 直接 ImportError。
"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from lrkit.pyfetch import _compatible_wheel  # noqa: E402

TAG = f"cp{sys.version_info.major}{sys.version_info.minor}"
IS64 = sys.maxsize > 2**32
ARCH = "win_amd64" if IS64 else "win32"


class CompatibleWheelTest(unittest.TestCase):
    def test_native_matching_abi_is_ok(self):
        ok, rank = _compatible_wheel(f"lxml-6.1.3-{TAG}-{TAG}-{ARCH}.whl")
        self.assertTrue(ok)
        self.assertGreater(rank, 0)

    def test_wrong_python_abi_rejected(self):
        ok, _ = _compatible_wheel("lxml-6.1.3-cp39-cp39-win_amd64.whl")
        self.assertFalse(ok if TAG != "cp39" else True)

    def test_abi3_from_older_python_is_ok(self):
        major, minor = sys.version_info.major, sys.version_info.minor
        if minor >= 3:
            tag = f"cp{major}{minor - 3}"
            ok, rank = _compatible_wheel(f"pymupdf-2.1-{tag}-abi3-{ARCH}.whl")
            self.assertTrue(ok)
            self.assertGreater(rank, 0)

    def test_pure_python_wheel_is_ok(self):
        ok, rank = _compatible_wheel("python_docx-1.2.0-py3-none-any.whl")
        self.assertTrue(ok)
        self.assertGreater(rank, 0)

    @unittest.skipUnless(IS64, "只在 64 位解释器上验证 32 位 wheel 被排除")
    def test_32bit_wheel_loses_to_64bit(self):
        ok32, rank32 = _compatible_wheel(f"lxml-6.1.3-{TAG}-{TAG}-win32.whl")
        ok64, rank64 = _compatible_wheel(f"lxml-6.1.3-{TAG}-{TAG}-win_amd64.whl")
        self.assertTrue(ok64)
        self.assertFalse(ok32)
        self.assertGreater(rank64, rank32)


if __name__ == "__main__":
    unittest.main()
