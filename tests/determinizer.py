import sys 
sys.path.insert(0, '')

from xl_utility.determinizer import *


def test_guessGender():
    assert guessGender() == "guess"

def test_generateUUID():
    assert generateUUID() == "genID"

def test_generateMockData():
    assert generateMockData() == "genMock"