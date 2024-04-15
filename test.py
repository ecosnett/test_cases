from Read_Excel import get_div, get_total

class tests(unittest.TestCase):
    def test_div(self):
        assert get_div(search) == str(ws["S"+str(count)].value)
        assert get_div(" ") == None

    def test_total(self):
        assert get_total() == str(ws["S9"].value)

if __name__=='__main__':
    unittest.main()
