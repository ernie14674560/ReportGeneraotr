import warnings

warnings.filterwarnings("ignore", "(?s).*MATPLOTLIBDATA.*", category=UserWarning)
import wm_app
import pickle
import sys

data = pickle.loads(sys.stdin.buffer.read())

if __name__ == '__main__':
    app = wm_app.WaferMapApp(**data)
