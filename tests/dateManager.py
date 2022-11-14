import sys 
sys.path.insert(0, '')

from xl_utility.dateManager import *


def test_getFullDate():
    assert getFullDate() == "getFull"

def test_getDateDetail():
    assert getDateDetail() == "getDetail"

def test_getShortDate():
    assert getShortDate() == "getShort"
