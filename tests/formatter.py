import sys 
sys.path.insert(0, '')

from xl_utility.formatter import *


def test_separateNames():
  assert separateNames() == "sepName"

def test_separateAddresses():
  assert separateAddresses() == "sepAdd"

def test_capitalizeFirstLetter():
  assert capitalizeFirstLetter() == "capFirst"

def test_capitalizeAll():
  assert capitalizeAll() == "capAll"