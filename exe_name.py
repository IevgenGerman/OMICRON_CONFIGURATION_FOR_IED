# -*- coding: utf-8 -*-
"""Prints the exe name with the version of ied2omicron.py: VERSION = "v 0.0.4 (dev)" -> IED2OMICRON_v0.0.4-dev
Used by build_exe.bat and .github/workflows/build.yml."""
import os, re
src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "ied2omicron.py"), encoding="utf-8").read()
v = re.search(r'^VERSION\s*=\s*"([^"]+)"', src, re.M).group(1)
v = re.sub(r"\s*\(([^)]*)\)", r"-\1", v)          # (dev) -> -dev
v = re.sub(r"[^\w.\-]+", "", v)                    # "v 0.0.4" -> v0.0.4
print("IED2OMICRON_" + v)
