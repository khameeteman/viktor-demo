###########
# The 'tests' folder contains all automated tests in one or more test_....py files. For more information on writing
# automated tests, see: https://docs.viktor.ai/docs/guides/best-practices/automated-testing
###########
import unittest

from viktor import Color

from app.controller import get_color


class TestGetColor(unittest.TestCase):
    def test_valid_value(self):
        color = get_color(50)
        self.assertEqual(color, Color(127, 255, 0))

    def test_invalid_value(self):
        with self.assertRaises(ValueError):
            get_color(101)  # value must be between 0 - 100!
