# coding: utf-8
import sys
from runner import run

if __name__ == '__main__':
    argv = sys.argv
    exit = run(argv)
    if exit:
        sys.exit(exit)
