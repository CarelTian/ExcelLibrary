import sys
import warnings
from ExcelLib import *
from unittest import TestCase
import builtins

class ExcelTest(TestCase):
    def setUp(self) :
        self.file1="./input/test1.xlsx"
        self.file2 = "./input/test2.xlsx"
        self.file3 = "./input/test3.xlsx"
        self.file4 = "./input/test4.csv"
        self.file5 = "./input/test5.csv"
        self.sheet="Sheet1"
        warnings.simplefilter('ignore', ResourceWarning)
    def test_1(self):
        ans1=[[0.9, 400, 83.7133815760754, 53.43], [0.9, 450, 86.4097027402944, 49.9], [1.68, 325, 19.6813590966625, 49.7], [1.68, 350, 36.801016971268, 47.21], [1.68, 300, 14.9688914921879, 46.94], [2.1, 400, 40.7971628999023, 42.04], [0.9, 400, 56.1163898081809, 41.42]]
        ans2=[[2, 1, 1, 0.49, 1.68, 275, 5.33483516226759, 2.55], [2, 1, 5, 1, 2.1, 250, 0.319963965933778, 2.19], [2, 1, 5, 1, 2.1, 300, 1.68467573268495, 2.17], [2, 1, 2, 1, 0.3, 250, 14.7871833611187, 1.96], [2, 1, 1, 0.49, 1.68, 250, 2.49268861247491, 1.89], [2, 1, 5, 1, 2.1, 275, 1.01545342802349, 1.65]]
        lt1=excel_to_list(self.file1,"E2:H8",assign=True)
        lt2=excel_to_list(self.file1,"A105:H105")
        self.assertEqual(ans1,lt1)
        self.assertEqual(ans2,lt2)
    def test_2(self):
        lt=[[1,2,3],[4,5,6],[7,8,9]]
        ret=list_to_excel(lt,self.file1,'J10:L10',overwrite=True)
        self.assertTrue(ret)
    def test_3(self):
        output="./input/merge.xlsx"
        self.assertTrue(merge_excel({self.file1:self.sheet,self.file2:'test1'},output))
    def test_4(self):
        ret=['标签分类', '装填方式', 'CO负载量', '装料比', '乙醇加入量', '温度', '乙醇转化率', 'C4选择性',None,None,None,None]
        self.assertEqual(ret,read_list(self.file1,'1'))
    def test_5(self):
        self.assertTrue(excel_replace(self.file1,1.68,666,'J27:L31'))

    def test_6(self):
        lt=[['标签分类', '装填方式', 'CO负载量', '装料比', '乙醇加入量', '温度', '乙醇转化率', 'C4选择性'], [1, 1, 1, 1, 0.9, 400, 83.7133815760754, 53.43], [1, 1, 1, 1, 2.1, 400, 40.7971628999023, 42.04], [1, 1, 1, 1, 0.9, 400, 56.1163898081809, 41.42], [1, 2, 1, 1, 1.68, 400, 43.595444389466, 41.08], [1, 1, 0.5, 1, 1.68, 400, 88.4393444439815, 41.02], [1, 2, 1, 1, 1.68, 400, 45.1352390825493, 38.7], [1, 1, 2, 1, 0.3, 400, 76.0198321376534, 38.23], [1, 2, 1, 1, 0.9, 400, 69.4, 38.17], [1, 1, 5, 1, 1.68, 400, 83.3476161758825, 37.33], [1, 1, 1, 1, 1.68, 400, 44.534966735128, 36.3], [1, 1, 1, 1, 0.3, 400, 76.0274161977392, 33.25], [1, 2, 1, 1, 1.68, 400, 63.2452382390403, 30.48], [1, 1, 1, 2.03, 1.68, 400, 40.047154305474, 27.91], [1, 2, 1, 1, 2.1, 400, 44.9818815507988, 25.83], [1, 1, 1, 0.49, 1.68, 400, 53.6152568624797, 22.3], [1, 2, 1, 1, 1.68, 400, 33.4895955281135, 21.45], [1, 2, 1, 1, 1.68, 400, 21.1001303126785, 21.21], [2, 1, 5, 1, 2.1, 400, 28.5943907570513, 10.29]]
        ret=excel_extract(self.file3,'F',bigger=370,smaller=410)
        self.assertEqual(lt,ret)
    def test_7(self):
        self.assertTrue(excel_pivot(self.file1,'A1:H1','CO负载量','max','./input/output1.xlsx'))
    def test_8(self):
        self.assertTrue(csv_split(self.file4,40,"./input/"))
        self.assertTrue(csv_split(self.file5,80000,"./input/",hasHead=True))
