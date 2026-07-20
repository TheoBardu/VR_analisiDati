#!/opt/anaconda3/bin/python3

"""Launcher della GUI. L'implementazione sta nel package GUI/."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from GUI.app import main

if __name__ == '__main__':
    main()
