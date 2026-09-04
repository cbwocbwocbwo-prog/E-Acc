import sys

from eacc_app.ui import run


if __name__ == "__main__":
    run(sys.argv[1] if len(sys.argv) > 1 else None)
