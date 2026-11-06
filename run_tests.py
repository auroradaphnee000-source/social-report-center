from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

root = Path(__file__).resolve().parent
sys.path.insert(0, str(root))
scratch = root / "_qa"
scratch.mkdir(exist_ok=True)
tempfile.tempdir = str(scratch)
suite = unittest.defaultTestLoader.discover(str(root / "tests"))
with patch("webapp.app.start_auth_refresher"):
    result = unittest.TextTestRunner(verbosity=2).run(suite)
sys.exit(not result.wasSuccessful())
