import tests.cellbase_test
import unittest


if __name__ == "__main__":
    suite = unittest.TestLoader().loadTestsFromModule(tests.cellbase_test)
    unittest.TextTestRunner().run(suite)