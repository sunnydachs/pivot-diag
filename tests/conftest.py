"""tests/ から fixture_builder（テスト用ヘルパー）を import 可能にする。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
