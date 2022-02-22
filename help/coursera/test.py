import unittest


class TestFactorize(unittest.TestCase):
    def test_wrong_types_raise_exception(self):
        for x in ['string', 1.5]:
            with self.subTest(x=x):
                self.assertRaises(TypeError, factorize, x)

    def test_negative(self):
        for x in [-1, -10, -100]:
            with self.subTest(x=x):
                self.assertRaises(ValueError, factorize, x)

    def test_zero_and_one_cases(self):
        res = {0: (0,), 1: (1,)}
        for x in res:
            with self.subTest(x=x):
                self.assertEqual(factorize(x), res[x])

    def test_simple_numbers(self):
        res = {3: (3,), 13: (13,), 29: (29,)}
        for x in res:
            with self.subTest(x=x):
                self.assertEqual(factorize(x), res[x])

    def test_two_simple_multipliers(self):
        res = {6: (2, 3), 26: (2, 13), 121: (11, 11)}
        for x in res:
            with self.subTest(x=x):
                self.assertEqual(factorize(x), res[x])

    def test_many_multipliers(self):
        res = {1001: (7, 11, 13), 9699690: (2, 3, 5, 7, 11, 13, 17, 19)}
        for x in res:
            with self.subTest(x=x):
                self.assertEqual(factorize(x), res[x])
