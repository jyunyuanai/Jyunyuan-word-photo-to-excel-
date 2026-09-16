# -*- coding: utf-8 -*-
"""資料夾照片 → 驗收相片 Excel"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _gallery_common import render_page

render_page("驗收相片")
