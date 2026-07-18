from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import main as controller_main


def main():
    path = controller_main.sync_regression_manifest()
    print(path)


if __name__ == "__main__":
    main()
