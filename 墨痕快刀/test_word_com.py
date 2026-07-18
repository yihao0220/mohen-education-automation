import sys
import os
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'core'))
from wps_helper import get_active_wps
wps = get_active_wps()
doc = wps.ActiveDocument
rng = doc.Paragraphs(1).Range
text = rng.Text
print(f"Start: {rng.Start}, End: {rng.End}, Length: {len(text)}, Delta: {rng.End - rng.Start}")
