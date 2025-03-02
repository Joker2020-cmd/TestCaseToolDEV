#导入库
import xmind
import test
#处理函数

def Rand_xmind(path):
     print('读取Xmind')
     workbook = xmind.load(path=path)
     sheet = workbook.getData()
     return sheet
test.rand1()


if __name__ == '__main__':
    print('xmind转换excel Testcase')
    Path="大模型应用_姓名_手机号.xmind"
    Rand_xmind(Path)
