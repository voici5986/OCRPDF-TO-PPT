# -*- coding: utf-8 -*-
"""
Check Python version for PaddlePaddle compatibility
"""
import sys

print("=" * 60)
print("Python Version Check")
print("=" * 60)

python_version = sys.version_info
print(f"\nYour Python version: {python_version.major}.{python_version.minor}.{python_version.micro}")

print("\nPaddlePaddle GPU support:")
print("  Python 3.8:  Supported")
print("  Python 3.9:  Supported")
print("  Python 3.10: Supported")
print("  Python 3.11: Supported")
print("  Python 3.12: Limited support")
print("  Python 3.13: NOT YET SUPPORTED")

if python_version.major == 3 and python_version.minor == 13:
    print("\n[!!] WARNING: Python 3.13 is too new!")
    print("    PaddlePaddle GPU does not support Python 3.13 yet.")
    print("\n    Solutions:")
    print("    1. Use Python 3.11 (recommended)")
    print("    2. Use Python 3.10")
    print("    3. Wait for PaddlePaddle to add Python 3.13 support")
    print("\n    Download Python 3.11:")
    print("    https://www.python.org/downloads/release/python-31110/")
elif python_version.major == 3 and python_version.minor >= 8:
    print(f"\n[OK] Python {python_version.major}.{python_version.minor} is supported!")
else:
    print(f"\n[!!] Python {python_version.major}.{python_version.minor} may not be supported")

print("=" * 60)
