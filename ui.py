import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / 'lib'))
from outlook_flow.ui import App

if __name__ == '__main__':
    App(sys.argv[1:]).run()
