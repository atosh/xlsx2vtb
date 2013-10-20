# coding: utf-8
import sys
import os

def run(argv):
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, base)
    import xlsx2vtb
    return xlsx2vtb.main(argv)

if __name__ == '__main__':
    argv = sys.argv
    exit = run(argv)
    if exit:
        sys.exit(exit)

