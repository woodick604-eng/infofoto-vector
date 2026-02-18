import sys
from pathlib import Path

app_py_path = Path("container/app.py")
b64_path = Path("container/static/logo_final_b64.txt")

with open(b64_path, "r") as f:
    real_b64 = f.read().strip()

content = app_py_path.read_text()

# Find and replace LOGO_PMA_B64
import re
new_content = re.sub(r'LOGO_PMA_B64 = """.*?"""', f'LOGO_PMA_B64 = """{real_b64}"""', content, flags=re.DOTALL)

# Ensure resize_to_box and to_jpeg_path are at the top
# Let's see if they are already there
if "def resize_to_box" not in new_content:
    print("Warning: resize_to_box not found in content, adding it.")
    # (The tool already saw it at line 111, so it should be there)

app_py_path.write_text(new_content)
print("Updated LOGO_PMA_B64 in app.py")
